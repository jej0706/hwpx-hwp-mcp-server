"""Pydantic models for MCP tool request/response payloads.

Keeping these in one place makes it easy to tell at a glance what shape the
LLM will see, and lets us evolve the schemas without touching every tool
module.
"""

from __future__ import annotations

from typing import Any, Dict, List, Literal, Optional

from pydantic import BaseModel, Field

SaveFormat = Literal["auto", "HWP", "HWPX", "PDF", "HTML", "DOCX", "TEXT"]
ExportFormat = Literal["text", "html", "pdf", "docx"]
AlignLiteral = Literal["left", "center", "right", "justify", "distribute"]


# ------------------------------------------------------------ Session models


class DocumentRef(BaseModel):
    doc_id: int = Field(..., description="0-based XHwpDocuments index")
    title: Optional[str] = None
    path: Optional[str] = None
    format: Optional[str] = None
    page_count: Optional[int] = None
    is_modified: Optional[bool] = None


class OpenResult(DocumentRef):
    pass


class SaveResult(BaseModel):
    saved: bool
    path: str
    format: Optional[str] = None
    backup_path: Optional[str] = None


class CloseResult(BaseModel):
    closed: bool
    doc_id: int


class ListDocumentsResult(BaseModel):
    documents: List[DocumentRef]


# ------------------------------------------------------------ Read models


class DocumentTextResult(BaseModel):
    text: str
    char_count: int
    page_count: Optional[int] = None


class FieldInfo(BaseModel):
    name: str
    index: int
    current_text: Optional[str] = None


class DocumentInfo(BaseModel):
    title: Optional[str] = None
    path: Optional[str] = None
    page_count: Optional[int] = None
    is_modified: Optional[bool] = None
    field_count: int = 0


class TableStructure(BaseModel):
    index: int
    rows: Optional[int] = None
    cols: Optional[int] = None
    caption: Optional[str] = None


class ImageStructure(BaseModel):
    index: int


class HeadingEntry(BaseModel):
    level: int
    text: str


class DocumentStructure(BaseModel):
    page_count: Optional[int] = None
    tables: List[TableStructure] = Field(default_factory=list)
    images: List[ImageStructure] = Field(default_factory=list)
    fields: List[FieldInfo] = Field(default_factory=list)
    headings: List[HeadingEntry] = Field(default_factory=list)


class SearchHit(BaseModel):
    match: str
    context: str


class SearchResult(BaseModel):
    query: str
    hit_count: int
    hits: List[SearchHit]


class ExportResult(BaseModel):
    exported: bool
    path: str
    format: ExportFormat


class TableCsvResult(BaseModel):
    table_index: int
    rows: int
    cols: int
    csv: str


# ------------------------------------------------------------ Template models


class FillFieldsResult(BaseModel):
    filled: int
    unknown_fields: List[str] = Field(default_factory=list)


class CreateFieldResult(BaseModel):
    created: bool
    name: str


class ReplaceTextResult(BaseModel):
    replaced: int


class FillTablePathResult(BaseModel):
    filled: int
    misses: List[str] = Field(default_factory=list)


# ------------------------------------------------------------ Create models


class InsertResult(BaseModel):
    inserted: bool
    detail: Optional[str] = None


class InsertTableResult(BaseModel):
    inserted: bool
    rows: int
    cols: int


class AppliedResult(BaseModel):
    applied: bool
    detail: Optional[str] = None


# ------------------------------------------------------------ Batch models


class BatchReplaceFileResult(BaseModel):
    path: str
    saved_as: str
    replaced: int
    ok: bool
    error: Optional[str] = None


class BatchReplaceResult(BaseModel):
    results: List[BatchReplaceFileResult]
    total_files: int
    total_replacements: int


class ConvertFileResult(BaseModel):
    src: str
    dst: str
    ok: bool
    error: Optional[str] = None


class ConvertResult(BaseModel):
    results: List[ConvertFileResult]
    total: int
    succeeded: int


# ------------------------------------------------------------ helpers


def to_dict(model: BaseModel) -> Dict[str, Any]:
    return model.model_dump(mode="json")
