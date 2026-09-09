from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class BoundingBox(BaseModel):
    """
    Axis-aligned bounding box in PDF points.

    Coordinate system:
        origin = top-left
        x increases to the right
        y increases downward
    """

    model_config = ConfigDict(extra="forbid")

    x0: float
    y0: float
    x1: float
    y1: float


class Point(BaseModel):
    model_config = ConfigDict(extra="forbid")

    x: float
    y: float


class TextSpan(BaseModel):
    """
    A single text span as reported by PyMuPDF.
    """

    model_config = ConfigDict(extra="forbid")

    text: str
    bbox: BoundingBox
    font: str | None = None
    size: float | None = None
    flags: int | None = None
    color: int | None = None
    ascender: float | None = None
    descender: float | None = None
    origin: Point | None = None


class DocumentElement(BaseModel):
    """
    Generic document element.

    This is intentionally domain-independent.
    """

    model_config = ConfigDict(extra="forbid")

    type: str

    text: str | None = None

    bbox: BoundingBox | None = None

    confidence: float | None = None

    source: str | None = None

    metadata: dict[str, Any] = Field(
        default_factory=dict
    )


class TextBlock(BaseModel):
    """
    A text block containing one or more lines/spans.
    """

    model_config = ConfigDict(extra="forbid")

    type: Literal["text"] = "text"

    bbox: BoundingBox

    number: int

    lines: list[dict[str, Any]] = Field(
        default_factory=list
    )

    text: str


class ImageAsset(BaseModel):
    """
    Reference to an image extracted from the PDF
    or a rendered page.
    """

    model_config = ConfigDict(extra="forbid")

    asset_id: str

    kind: Literal[
        "embedded",
        "page_render",
    ]

    page_number: int

    path: str

    width: int | None = None

    height: int | None = None

    extension: str | None = None

    xref: int | None = None

    bbox: BoundingBox | None = None

    source: str


class DrawingElement(BaseModel):
    """
    Vector drawing information preserved from the PDF.

    Engineering drawings frequently contain important
    graphical information that is not represented as
    ordinary text.
    """

    model_config = ConfigDict(
        extra="allow"
    )

    type: Literal["drawing"] = "drawing"

    bbox: BoundingBox

    seqno: int | None = None

    layer: str | None = None

    width: float | None = None

    color: Any = None

    fill: Any = None

    opacity: float | None = None

    fill_opacity: float | None = None

    close_path: bool | None = None

    items: list[Any] = Field(
        default_factory=list
    )


class PageInfo(BaseModel):

    model_config = ConfigDict(extra="forbid")

    page_number: int

    width: float

    height: float

    rotation: int

    mediabox: BoundingBox

    cropbox: BoundingBox

    text_char_count: int

    word_count: int

    text_block_count: int

    embedded_image_count: int

    drawing_count: int

    has_native_text: bool

    needs_ocr_candidate: bool

    elements: list[dict[str, Any]] = Field(
        default_factory=list
    )

    images: list[ImageAsset] = Field(
        default_factory=list
    )

    drawings: list[DrawingElement] = Field(
        default_factory=list
    )


class DocumentInfo(BaseModel):

    model_config = ConfigDict(extra="forbid")

    file_name: str

    source_path: str

    file_size_bytes: int

    sha256: str

    page_count: int

    pdf_metadata: dict[str, Any] = Field(
        default_factory=dict
    )

    extraction_engine: str

    extraction_engine_version: str

    extracted_at_utc: str


class RawDocument(BaseModel):
    """
    Stage-1 universal document representation.

    IMPORTANT:
    No domain-specific fields are defined here.

    For example, this model does NOT know about:

        part_number
        nominal
        measured
        tolerance
        balloon_number
        pass_fail

    Those will be introduced in later semantic
    understanding stages.
    """

    schema_name: str = (
        "pdf-intelligence.raw-document"
    )

    schema_version: str = "1.0.0"

    document: DocumentInfo

    pages: list[PageInfo]