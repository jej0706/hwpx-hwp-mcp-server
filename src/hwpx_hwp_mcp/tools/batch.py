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
