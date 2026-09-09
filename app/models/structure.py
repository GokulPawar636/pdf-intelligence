from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class Region(BaseModel):
    """
    A visually/spatially detected region of a PDF page.

    No business meaning is assigned here.
    """

    model_config = ConfigDict(extra="forbid")

    region_id: str

    page_number: int

    region_type: Literal[
        "text_region",
        "table_candidate",
        "graphic_region",
    ]

    bbox: dict[str, float]

    element_indexes: list[int] = Field(
        default_factory=list
    )

    text: str = ""

    confidence: float = 0.0

    evidence: dict[str, Any] = Field(
        default_factory=dict
    )


class TableCell(BaseModel):
    """
    A dynamically detected table cell.
    """

    model_config = ConfigDict(extra="forbid")

    row: int

    column: int

    text: str = ""

    bbox: dict[str, float] | None = None

    source_word_indexes: list[int] = Field(
        default_factory=list
    )


class TableCandidate(BaseModel):
    """
    A region that geometrically resembles a table.

    Important:
    It is called TableCandidate because Stage 2 has not
    semantically proven that it is a table.
    """

    model_config = ConfigDict(extra="forbid")

    table_id: str

    page_number: int

    bbox: dict[str, float]

    row_count: int

    column_count: int

    cells: list[TableCell] = Field(
        default_factory=list
    )

    confidence: float = 0.0

    evidence: dict[str, Any] = Field(
        default_factory=dict
    )


class PageStructure(BaseModel):

    warnings: list[str] = Field(default_factory=list)

    model_config = ConfigDict(extra="forbid")

    page_number: int

    regions: list[Region] = Field(
        default_factory=list
    )

    tables: list[TableCandidate] = Field(
        default_factory=list
    )


class StructuredDocument(BaseModel):
    """
    Stage-2 representation.

    Converts raw PDF information into spatial structure,
    but does not assign business meaning.
    """

    model_config = ConfigDict(extra="forbid")

    schema_name: str = (
        "pdf-intelligence.structured-document"
    )

    schema_version: str = "2.0.0"

    source_schema: str

    source_schema_version: str

    document: dict[str, Any]

    pages: list[PageStructure]