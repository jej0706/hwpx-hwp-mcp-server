"""Batch / bulk tools (category E).

These tools process many files in a single request. Because every call
lands on the dedicated COM worker thread, files are processed sequentially
and sharing a single Hancom instance — no risk of races.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import (
    BatchReplaceFileResult,
    BatchReplaceResult,
    ConvertFileResult,
    ConvertResult,
    to_dict,
)
from ..utils.paths import (
    backup_file,
    ensure_abs_windows_path,
    ensure_output_path,
    iter_input_files,
    resolve_save_format,
)


def _close_active(hwp: Any) -> None:
    """Close whatever document is currently active, discarding changes."""
    try:
        doc = hwp.XHwpDocuments.Item(hwp.XHwpDocuments.Count - 1)
        try:
            doc.Close(isDirty=False)
        except TypeError:
            doc.Close(False)
    except Exception:  # noqa: BLE001
        pass


def register(mcp: FastMCP) -> None:
    @mcp.tool(
        description=(
            "Run a list of find/replace pairs across many files. Each file "
            "is opened, every replacement is applied, then saved (to the "
            "original path, or to output_dir if given). Good for bulk "
            "templating like 'update 2025 → 2026 in every contract'."
        ),
    )
    async def batch_replace_in_files(
        replacements: List[Dict[str, str]] = Field(
            ...,
            description="List of {old, new} objects. Applied in order to each file.",
        ),
        input_paths: Optional[List[str]] = Field(
            None,
            description="Explicit list of files. Alternatively use folder+glob.",
        ),
        folder: Optional[str] = Field(
            None,
            description="Folder to scan. Used only if input_paths is omitted.",
        ),
        glob: str = Field("*.hwp*", description="Glob pattern when scanning folder"),
        output_dir: Optional[str] = Field(
            None,
            description="If set, save results into this folder with the same filename. "
            "Otherwise files are edited in place (with .bak backup).",
        ),
        backup: bool = Field(True, description="Create .bak backups when overwriting in place"),
    ) -> dict:
        files = iter_input_files(input_paths, folder=folder, glob=glob)
        out_dir: Optional[Path] = None
        if output_dir:
            out_dir = ensure_abs_windows_path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)

        # Normalize replacements once, outside the COM worker.
        norm: list[tuple[str, str]] = []
        for item in replacements:
            old = str(item.get("old", ""))
            new = str(item.get("new", ""))
            if not old:
                raise HwpError("replacements 항목에 빈 'old' 가 있습니다.")
            norm.append((old, new))

        def _do(hwp: Any) -> BatchReplaceResult:
            results: list[BatchReplaceFileResult] = []
            total_replacements = 0
            for src in files:
                saved_as: str = str(src)
                count = 0
                ok = True
                err: Optional[str] = None
                try:
                    if not hwp.open(str(src), format="", arg="lock:true"):
                        raise HwpError("open returned False")
                    for old, new in norm:
                        res = hwp.find_replace_all(old, new, regex=False)
                        try:
                            count += int(res)
                        except Exception:  # noqa: BLE001
                            count += int(bool(res))

                    if out_dir is not None:
                        dst = out_dir / src.name
                        hwp.save_as(str(dst), format=resolve_save_format("auto", dst) or "HWP")
                        saved_as = str(dst)
                    else:
                        if backup:
                            try:
                                backup_file(src)
                            except Exception:  # noqa: BLE001
                                pass
                        hwp.save(save_if_dirty=True)
                        saved_as = str(src)
                except Exception as exc:  # noqa: BLE001
                    ok = False
                    err = str(exc)
                finally:
                    _close_active(hwp)

                results.append(
                    BatchReplaceFileResult(
                        path=str(src),
                        saved_as=saved_as,
                        replaced=count,
                        ok=ok,
                        error=err,
                    )
                )
                if ok:
                    total_replacements += count
            return BatchReplaceResult(
                results=results,
                total_files=len(files),
                total_replacements=total_replacements,
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "Convert a list of files to a target format (hwp / hwpx / pdf / "
            "html / docx). Output files land in output_dir with the same "
            "base name but the new extension."
        ),
    )
    async def convert_files(
        input_paths: List[str] = Field(..., description="Files to convert"),
        target_format: str = Field(..., description="hwp | hwpx | pdf | html | docx"),
        output_dir: str = Field(..., description="Destination folder (created if missing)"),
    ) -> dict:
        fmt = target_format.lower()
        ext_map = {
            "hwp": ".hwp",
            "hwpx": ".hwpx",
            "pdf": ".pdf",
            "html": ".html",
            "docx": ".docx",
        }
        if fmt not in ext_map:
            raise HwpError(f"target_format 이 유효하지 않습니다: {target_format!r}")

        files = iter_input_files(input_paths)
        out_dir = ensure_abs_windows_path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        def _do(hwp: Any) -> ConvertResult:
            results: list[ConvertFileResult] = []
            succeeded = 0
            for src in files:
                dst = out_dir / (src.stem + ext_map[fmt])
                ok = True
                err: Optional[str] = None
                try:
                    if not hwp.open(str(src), format="", arg="lock:true"):
                        raise HwpError("open returned False")
                    hwp.save_as(str(dst), format=resolve_save_format(fmt, dst) or "HWP")
                except Exception as exc:  # noqa: BLE001
                    ok = False
                    err = str(exc)
                finally:
                    _close_active(hwp)
                results.append(
                    ConvertFileResult(src=str(src), dst=str(dst), ok=ok, error=err)
                )
                if ok:
                    succeeded += 1
            return ConvertResult(results=results, total=len(files), succeeded=succeeded)

        return to_dict(await session.call(_do))

    # ------------------------------------------------------------------ NEW

    @mcp.tool(
        description=(
            "여러 HWP 파일 합치기 / Merge multiple HWP/HWPX files into a "
            "single output document. Files are concatenated in the order "
            "given. A page break is inserted between each source document "
            "when page_break_between=True (default).\n\n"
            "output_format: 'hwp' | 'hwpx' | 'auto' (inferred from "
            "output_path extension)."
        ),
    )
    async def merge_documents(
        input_paths: List[str] = Field(
            ...,
            description="Ordered list of absolute HWP/HWPX file paths to merge",
        ),
        output_path: str = Field(
            ...,
            description="Absolute path for the merged output file (.hwp or .hwpx)",
        ),
        page_break_between: bool = Field(
            True,
            description="Insert a page break between each source document (default True)",
        ),
        output_format: str = Field(
            "auto",
            description="Output format: auto | hwp | hwpx (auto = from extension)",
        ),
    ) -> dict:
        files = iter_input_files(input_paths)
        if not files:
            raise HwpError("input_paths 가 비어 있습니다.")
        out = ensure_output_path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)

        def _do(hwp: Any) -> dict:
            results: list[dict] = []
            total_ok = 0

            # Open the first document as the base
            first = files[0]
            if not hwp.open(str(first), format="", arg="lock:true"):
                raise HwpError(f"첫 번째 파일을 열 수 없습니다: {first}")
            results.append({"path": str(first), "ok": True})
            total_ok += 1

            # Append subsequent documents
            for src in files[1:]:
                ok = True
                err: Optional[str] = None
                try:
                    # Move caret to end of current document
                    hwp.HAction.Run("MoveDocEnd")
                    if page_break_between:
                        hwp.BreakPage()

                    # Open the next file in a new tab, copy all content
                    if not hwp.open(str(src), format="", arg="lock:true"):
                        raise HwpError(f"파일을 열 수 없습니다: {src}")
                    hwp.SelectAll()
                    hwp.HAction.Run("Copy")

                    # Switch back to the base document
                    doc_count = int(hwp.XHwpDocuments.Count)
                    hwp.switch_to(doc_count - 2)  # base = second-to-last
                    hwp.HAction.Run("MoveDocEnd")
                    hwp.HAction.Run("Paste")

                    # Close the just-copied document without saving
                    hwp.switch_to(doc_count - 1)
                    try:
                        hwp.XHwpDocuments.Item(doc_count - 1).Close(False)
                    except Exception:  # noqa: BLE001
                        try:
                            hwp.XHwpDocuments.Item(doc_count - 1).Close(
                                isDirty=False
                            )
                        except Exception:  # noqa: BLE001
                            _close_active(hwp)

                    hwp.switch_to(0)
                    total_ok += 1

                except Exception as exc:  # noqa: BLE001
                    ok = False
                    err = str(exc)

                results.append({"path": str(src), "ok": ok, "error": err})

            # Save the merged document
            fmt = resolve_save_format(output_format, out) or "HWP"
            hwp.save_as(str(out), format=fmt)

            return {
                "merged": str(out),
                "total_files": len(files),
                "succeeded": total_ok,
                "results": results,
            }

        return await session.call(_do)

    @mcp.tool(
        description=(
            "두 문서 텍스트 비교 / Compare the plain text of two HWP/HWPX "
            "documents and return a unified diff. Formatting differences are "
            "ignored — only text content is compared. context_lines controls "
            "how many unchanged lines appear around each change block."
        ),
    )
    async def compare_documents(
        path1: str = Field(
            ..., description="Absolute path to the first (reference) document"
        ),
        path2: str = Field(
            ..., description="Absolute path to the second (modified) document"
        ),
        context_lines: int = Field(
            2,
            description="Unchanged context lines around each change block (0–10)",
            ge=0,
            le=10,
        ),
    ) -> dict:
        from ..utils.paths import ensure_existing_file

        p1 = ensure_existing_file(path1)
        p2 = ensure_existing_file(path2)

        def _do(hwp: Any) -> dict:
            texts: list[str] = []
            for src in (p1, p2):
                if not hwp.open(str(src), format="", arg="lock:true"):
                    raise HwpError(f"파일을 열 수 없습니다: {src}")
                try:
                    t = hwp.get_text_file("TEXT", "")
                except Exception:  # noqa: BLE001
                    t = ""
                texts.append(t or "")
                try:
                    cnt = int(hwp.XHwpDocuments.Count)
                    hwp.XHwpDocuments.Item(cnt - 1).Close(isDirty=False)
                except Exception:  # noqa: BLE001
                    pass

            import difflib

            lines1 = texts[0].splitlines(keepends=True)
            lines2 = texts[1].splitlines(keepends=True)
            diff = list(
                difflib.unified_diff(
                    lines1,
                    lines2,
                    fromfile=p1.name,
                    tofile=p2.name,
                    n=int(context_lines),
                )
            )
            changed = sum(
                1 for ln in diff if ln.startswith("+") or ln.startswith("-")
            )
            return {
                "identical": len(diff) == 0,
                "changed_lines": changed,
                "diff": "".join(diff)[:8000],
                "path1": str(p1),
                "path2": str(p2),
            }

        return await session.call(_do)

    @mcp.tool(
        description=(
            "여러 파일 누름틀 일괄 채우기 / Fill named fields (누름틀) with "
            "the same values across many HWP/HWPX files. Each file is opened, "
            "put_field_text is called, then saved. Useful for mass "
            "certificate / contract generation from a single template."
        ),
    )
    async def batch_fill_fields(
        input_paths: List[str] = Field(
            ..., description="List of absolute HWP/HWPX file paths"
        ),
        field_values: Dict[str, str] = Field(
            ...,
            description=(
                "Field name → value mapping applied to every file. "
                "Use {{0}}/{{1}} suffixes for multiple instances of the "
                "same field."
            ),
        ),
        output_dir: Optional[str] = Field(
            None,
            description=(
                "If set, save results here. Otherwise overwrite in place "
                "(with .bak backup)."
            ),
        ),
        backup: bool = Field(
            True, description="Create .bak backup when saving in place"
        ),
    ) -> dict:
        files = iter_input_files(input_paths)
        if not files:
            raise HwpError("input_paths 가 비어 있습니다.")
        if not field_values:
            raise HwpError("field_values 가 비어 있습니다.")

        out_dir: Optional[Path] = None
        if output_dir:
            out_dir = ensure_abs_windows_path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)

        def _do(hwp: Any) -> dict:
            results: list[dict] = []
            succeeded = 0

            for src in files:
                ok = True
                err: Optional[str] = None
                saved_as = str(src)
                try:
                    if not hwp.open(str(src), format="", arg="lock:true"):
                        raise HwpError("open returned False")
                    hwp.put_field_text(field_values)

                    if out_dir is not None:
                        dst = out_dir / src.name
                        hwp.save_as(
                            str(dst),
                            format=resolve_save_format("auto", dst) or "HWP",
                        )
                        saved_as = str(dst)
                    else:
                        if backup:
                            try:
                                backup_file(src)
                            except Exception:  # noqa: BLE001
                                pass
                        hwp.save(save_if_dirty=True)

                    succeeded += 1
                except Exception as exc:  # noqa: BLE001
                    ok = False
                    err = str(exc)
                finally:
                    try:
                        cnt = int(hwp.XHwpDocuments.Count)
                        hwp.XHwpDocuments.Item(cnt - 1).Close(isDirty=False)
                    except Exception:  # noqa: BLE001
                        pass

                results.append(
                    {
                        "path": str(src),
                        "saved_as": saved_as,
                        "ok": ok,
                        "error": err,
                    }
                )

            return {
                "results": results,
                "total_files": len(files),
                "succeeded": succeeded,
                "fields_applied": list(field_values.keys()),
            }

        return await session.call(_do)
