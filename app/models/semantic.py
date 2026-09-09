from __future__ import annotations

from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class BoundingBox(BaseModel):
    x0: float
    y0: float
    x1: float
    y1: float


class SourceReference(BaseModel):
    page_number: int

    source_text: Optional[str] = None

    bbox: Optional[BoundingBox] = None

    extraction_method: str = "unknown"

    confidence: float = 0.0

    source_id: Optional[str] = None


class SemanticValue(BaseModel):
    raw_value: Optional[str] = None

    normalized_value: Optional[Any] = None

    unit: Optional[str] = None

    source: Optional[SourceReference] = None


class SemanticField(BaseModel):
    name: str

    value: SemanticValue


class SemanticRecord(BaseModel):
    record_id: Optional[str] = None

    record_type: str

    fields: List[SemanticField] = Field(
        default_factory=list
    )

    source_pages: List[int] = Field(
        default_factory=list
    )

    confidence: float = 0.0

    related_record_ids: List[str] = Field(
        default_factory=list
    )

    warnings: List[str] = Field(
        default_factory=list
    )


class SemanticDocument(BaseModel):
    schema_name: str = (
        "pdf-intelligence.semantic-document"
    )

    schema_version: str = "3.1.0"

    source_schema: str

    source_schema_version: str

    document_id: str

    document_metadata: Dict[str, Any] = Field(
        default_factory=dict
    )

    records: List[SemanticRecord] = Field(
        default_factory=list
    )

    global_notes: List[str] = Field(
        default_factory=list
    )

    extraction_model: str

    extraction_provider: str = "groq"

    extraction_confidence: float = 0.0

    warnings: List[str] = Field(
        default_factory=list
    )

    raw_model_response: List[Dict[str, Any]] = Field(
        default_factory=list
    )