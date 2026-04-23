"""Output and security tools (category I).

print_document, protect_document, get_page_as_image.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, List, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from ..backend.errors import HwpError
from ..backend.hancom_com import session
from ..models import AppliedResult, InsertResult, to_dict
from ..utils.paths import ensure_abs_windows_path
from .session import _require_doc


def register(mcp: FastMCP) -> None:

    @mcp.tool(
        description=(
            "문서 인쇄 / Print the document to a printer. "
            "If printer_name is omitted the default system printer is used. "
            "page_range examples: '' (all), '1-3', '1,3,5-7'. "
            "copies sets the number of copies (default 1).\n\n"
            "⚠️ This sends directly to the spooler — no preview dialog."
        ),
    )
    async def print_document(
        doc_id: int = Field(..., description="Document index from open_document"),
        copies: int = Field(1, description="Number of copies", ge=1, le=99),
        page_range: str = Field(
            "",
            description=(
                "Pages to print. '' = all, '1-3' = pages 1-3, "
                "'1,3,5-7' = pages 1, 3, and 5-7."
            ),
        ),
        printer_name: Optional[str] = Field(
            None, description="Printer name; None = system default printer"
        ),
        collate: bool = Field(True, description="Collate copies"),
        duplex: bool = Field(False, description="Print duplex (both sides)"),
    ) -> dict:
        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                pset = hwp.HParameterSet.HPrint
                hwp.HAction.GetDefault("FilePrint", pset.HSet)

                if printer_name:
                    pset.PrinterName = str(printer_name)
                pset.Collate = 1 if collate else 0
                pset.NumCopy = int(copies)  # NOTE: NumCopy (not NumCopies)

                # PrintMethod: 0=all, 2=custom range (stored in Range field)
                if page_range.strip():
                    pset.PrintMethod = 2  # custom range
                    # Range uses a custom format; RangeCustom holds the string
                    try:
                        pset.Range = str(page_range)
                    except Exception:  # noqa: BLE001
                        try:
                            pset.RangeCustom = str(page_range)
                        except Exception:  # noqa: BLE001
                            pass
                else:
                    pset.PrintMethod = 0  # all pages

                ok = bool(hwp.HAction.Execute("FilePrint", pset.HSet))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"인쇄 실패: {exc}") from exc
            printer_info = printer_name or "기본 프린터"
            return AppliedResult(
                applied=ok,
                detail=(
                    f"printer={printer_info!r} copies={copies} "
                    f"range={page_range or 'all'}"
                ),
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "문서 보호 설정 / Protect or unprotect the document with a "
            "password. protect_type controls what is protected:\n"
            "- 'all' — entire document (읽기 전용 포함)\n"
            "- 'edit' — prevent editing but allow reading\n"
            "- 'none' — remove protection (password required to unprotect)\n\n"
            "⚠️ If you forget the password, the document cannot be "
            "unprotected without third-party tools."
        ),
    )
    async def protect_document(
        doc_id: int = Field(..., description="Document index from open_document"),
        password: str = Field(..., description="Password for protection/unprotection"),
        protect_type: str = Field(
            "edit",
            description="Protection type: all | edit | none (none = remove protection)",
        ),
    ) -> dict:
        ptype = protect_type.strip().lower()
        if ptype not in {"all", "edit", "none"}:
            raise HwpError(
                f"protect_type 이 유효하지 않습니다: {protect_type!r} (all / edit / none)"
            )

        def _do(hwp: Any) -> AppliedResult:
            _require_doc(hwp, doc_id)
            try:
                # HFileSecurity is the correct parameter set for document protection
                pset = hwp.HParameterSet.HFileSecurity
                hwp.HAction.GetDefault("FileSecurity", pset.HSet)
                pset.PasswordString = str(password)
                if ptype == "none":
                    # Remove protection: set empty password
                    pset.PasswordString = ""
                    pset.PasswordFullRange = 0
                    pset.NoCopy = 0
                    pset.NoPrint = 0
                elif ptype == "edit":
                    pset.PasswordFullRange = 1  # full document range
                    pset.PasswordAsk = 1
                    pset.NoCopy = 0
                    pset.NoPrint = 0
                else:  # all
                    pset.PasswordFullRange = 1
                    pset.PasswordAsk = 1
                    pset.NoCopy = 1
                    pset.NoPrint = 1
                ok = bool(hwp.HAction.Execute("FileSecurity", pset.HSet))
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"문서 보호 설정 실패: {exc}") from exc
            return AppliedResult(
                applied=ok,
                detail=f"protect_type={protect_type}",
            )

        return to_dict(await session.call(_do))

    @mcp.tool(
        description=(
            "페이지를 이미지로 저장 / Render one or all pages of the document "
            "as image files and save them to output_path.\n\n"
            "- page=0 → current page only\n"
            "- page=N (1-based) → that specific page\n"
            "- page=-1 → all pages (output_path becomes a pattern: "
            "'img.png' → 'img001.png', 'img002.png', …)\n\n"
            "format: 'png' | 'jpg' | 'bmp' (default 'png')\n"
            "resolution: DPI, default 150 (use 300 for print quality)"
        ),
    )
    async def get_page_as_image(
        doc_id: int = Field(..., description="Document index from open_document"),
        output_path: str = Field(
            ...,
            description=(
                "Absolute Windows path for the output image. "
                "For multi-page (page=-1) the base name gets a 3-digit "
                "suffix per page."
            ),
        ),
        page: int = Field(
            0,
            description=(
                "Page to render: 0 = current page, N = page N (1-based), "
                "-1 = all pages"
            ),
        ),
        format: str = Field(
            "png", description="Image format: png | jpg | bmp"
        ),
        resolution: int = Field(
            150, description="Resolution in DPI (72–1200)", ge=72, le=1200
        ),
    ) -> dict:
        fmt = format.strip().lower()
        if fmt == "jpeg":
            fmt = "jpg"
        if fmt not in {"png", "jpg", "bmp"}:
            raise HwpError(
                f"format 이 유효하지 않습니다: {format!r} (png / jpg / bmp)"
            )

        out_path = ensure_abs_windows_path(output_path)
        # Ensure parent directory exists
        out_path.parent.mkdir(parents=True, exist_ok=True)

        def _do(hwp: Any) -> dict:
            _require_doc(hwp, doc_id)
            try:
                page_count = int(hwp.PageCount)
                if page > page_count:
                    raise HwpError(
                        f"page={page} 가 문서 총 페이지 수 {page_count}를 초과합니다."
                    )
                ok = bool(
                    hwp.create_page_image(
                        path=str(out_path),
                        pgno=int(page),
                        resolution=int(resolution),
                        depth=24,
                        format="bmp",  # pyhwpx renders BMP then converts
                    )
                )
            except HwpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise HwpError(f"페이지 이미지 저장 실패: {exc}") from exc

            # Collect actually created files
            if page == -1:
                base = out_path.stem
                ext = out_path.suffix
                parent = out_path.parent
                saved = sorted(str(p) for p in parent.glob(f"{base}*{ext}"))
            else:
                saved = [str(out_path)] if ok else []

            return {
                "saved": ok,
                "files": saved,
                "page": page,
                "resolution": resolution,
                "format": fmt,
            }

        return await session.call(_do)
