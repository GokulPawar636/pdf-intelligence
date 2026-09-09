from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import fitz

from app.models.document import (
    BoundingBox,
    DrawingElement,
    ImageAsset,
    PageInfo,
    Point,
    RawDocument,
)


class PDFLoader:
    """
    Load a PDF into a layout-preserving,
    domain-agnostic representation.
    """

    OCR_CHAR_THRESHOLD = 20

    def __init__(
        self,
        pdf_path: str | Path,
        output_dir: str | Path,
        *,
        render_pages: bool = True,
        extract_embedded_images: bool = True,
        include_drawings: bool = True,
        render_dpi: int = 150,
    ) -> None:

        self.pdf_path = (
            Path(pdf_path)
            .expanduser()
            .resolve()
        )

        self.output_dir = (
            Path(output_dir)
            .expanduser()
            .resolve()
        )

        self.render_pages = render_pages

        self.extract_embedded_images = (
            extract_embedded_images
        )

        self.include_drawings = (
            include_drawings
        )

        self.render_dpi = render_dpi

        if not self.pdf_path.exists():
            raise FileNotFoundError(
                f"PDF not found: {self.pdf_path}"
            )

        if self.pdf_path.suffix.lower() != ".pdf":
            raise ValueError(
                f"Expected a PDF file, "
                f"got: {self.pdf_path.suffix}"
            )

        if render_dpi <= 0:
            raise ValueError(
                "render_dpi must be greater than zero"
            )

    @staticmethod
    def _bbox(
        rect: fitz.Rect,
    ) -> BoundingBox:

        return BoundingBox(
            x0=round(
                float(rect.x0),
                4,
            ),
            y0=round(
                float(rect.y0),
                4,
            ),
            x1=round(
                float(rect.x1),
                4,
            ),
            y1=round(
                float(rect.y1),
                4,
            ),
        )

    @staticmethod
    def _point(
        point: tuple[float, float] | None,
    ) -> Point | None:

        if point is None:
            return None

        return Point(
            x=round(
                float(point[0]),
                4,
            ),
            y=round(
                float(point[1]),
                4,
            ),
        )

    @staticmethod
    def _json_safe(
        value: Any,
    ) -> Any:
        """
        Convert PyMuPDF objects into JSON-compatible
        Python primitives.
        """

        if isinstance(
            value,
            fitz.Rect,
        ):
            return [
                round(float(value.x0), 4),
                round(float(value.y0), 4),
                round(float(value.x1), 4),
                round(float(value.y1), 4),
            ]

        if isinstance(
            value,
            fitz.Point,
        ):
            return [
                round(float(value.x), 4),
                round(float(value.y), 4),
            ]

        if isinstance(
            value,
            (tuple, list),
        ):
            return [
                PDFLoader._json_safe(item)
                for item in value
            ]

        if isinstance(
            value,
            dict,
        ):
            return {
                str(key):
                    PDFLoader._json_safe(value)
                for key, value in value.items()
            }

        if isinstance(
            value,
            (
                str,
                int,
                float,
                bool,
            ),
        ) or value is None:

            return value

        return str(value)

    def _sha256(self) -> str:

        digest = hashlib.sha256()

        with self.pdf_path.open(
            "rb"
        ) as file:

            for chunk in iter(
                lambda: file.read(
                    1024 * 1024
                ),
                b"",
            ):

                digest.update(chunk)

        return digest.hexdigest()

    def _extract_text_elements(
        self,
        page: fitz.Page,
    ) -> tuple[
        list[dict[str, Any]],
        int,
        int,
    ]:
        """
        Extract blocks, lines, spans and words.

        No semantic interpretation happens here.
        """

        text_dict = page.get_text(
            "dict",
            sort=False,
        )

        words_raw = page.get_text(
            "words",
            sort=False,
        )

        elements: list[
            dict[str, Any]
        ] = []

        block_count = 0

        for (
            block_number,
            block,
        ) in enumerate(
            text_dict.get(
                "blocks",
                [],
            ),
            start=1,
        ):

            block_type = block.get(
                "type"
            )

            # PyMuPDF:
            # 0 = text
            # 1 = image
            #
            # Images are handled separately.
            if block_type != 0:
                continue

            block_count += 1

            lines_out = []

            for (
                line_number,
                line,
            ) in enumerate(
                block.get(
                    "lines",
                    [],
                ),
                start=1,
            ):

                spans_out = []

                line_text_parts = []

                for (
                    span_number,
                    span,
                ) in enumerate(
                    line.get(
                        "spans",
                        [],
                    ),
                    start=1,
                ):

                    text = span.get(
                        "text",
                        "",
                    )

                    line_text_parts.append(
                        text
                    )

                    span_out = {
                        "span_number":
                            span_number,

                        "text":
                            text,

                        "bbox":
                            self._bbox(
                                fitz.Rect(
                                    span["bbox"]
                                )
                            ).model_dump(),

                        "font":
                            span.get(
                                "font"
                            ),

                        "size":
                            (
                                round(
                                    float(
                                        span["size"]
                                    ),
                                    4,
                                )
                                if span.get(
                                    "size"
                                ) is not None
                                else None
                            ),

                        "flags":
                            span.get(
                                "flags"
                            ),

                        "color":
                            span.get(
                                "color"
                            ),

                        "ascender":
                            span.get(
                                "ascender"
                            ),

                        "descender":
                            span.get(
                                "descender"
                            ),

                        "origin":
                            (
                                self._point(
                                    span.get(
                                        "origin"
                                    )
                                ).model_dump()
                                if span.get(
                                    "origin"
                                ) is not None
                                else None
                            ),
                    }

                    spans_out.append(
                        span_out
                    )

                lines_out.append(
                    {
                        "line_number":
                            line_number,

                        "bbox":
                            self._bbox(
                                fitz.Rect(
                                    line["bbox"]
                                )
                            ).model_dump(),

                        "text":
                            "".join(
                                line_text_parts
                            ),

                        "spans":
                            spans_out,

                        "wmode":
                            line.get(
                                "wmode"
                            ),

                        "dir":
                            self._json_safe(
                                line.get(
                                    "dir"
                                )
                            ),
                    }
                )

            block_text = "\n".join(
                line["text"]
                for line in lines_out
            )

            elements.append(
                {
                    "type": "text",

                    "block_number":
                        block_number,

                    "bbox":
                        self._bbox(
                            fitz.Rect(
                                block["bbox"]
                            )
                        ).model_dump(),

                    "text":
                        block_text,

                    "lines":
                        lines_out,
                }
            )

        # Words are also preserved independently.
        #
        # This is important later when we need
        # spatial relationships between words.
        for word in words_raw:

            if len(word) < 8:
                continue

            elements.append(
                {
                    "type": "word",

                    "text":
                        str(word[4]),

                    "bbox":
                        self._bbox(
                            fitz.Rect(
                                word[:4]
                            )
                        ).model_dump(),

                    "block_number":
                        int(word[5]),

                    "line_number":
                        int(word[6]),

                    "word_number":
                        int(word[7]),
                }
            )

        return (
            elements,
            len(words_raw),
            block_count,
        )

    def _extract_drawings(
        self,
        page: fitz.Page,
    ) -> list[DrawingElement]:

        if not self.include_drawings:
            return []

        drawings = []

        for drawing in page.get_drawings():

            rect = drawing.get(
                "rect"
            )

            if rect is None:
                continue

            item_list = []

            for item in drawing.get(
                "items",
                [],
            ):

                item_list.append(
                    self._json_safe(
                        item
                    )
                )

            drawings.append(
                DrawingElement(
                    bbox=self._bbox(
                        rect
                    ),

                    seqno=drawing.get(
                        "seqno"
                    ),

                    layer=drawing.get(
                        "layer"
                    ),

                    width=drawing.get(
                        "width"
                    ),

                    color=self._json_safe(
                        drawing.get(
                            "color"
                        )
                    ),

                    fill=self._json_safe(
                        drawing.get(
                            "fill"
                        )
                    ),

                    opacity=drawing.get(
                        "stroke_opacity",
                        drawing.get(
                            "opacity"
                        ),
                    ),

                    fill_opacity=drawing.get(
                        "fill_opacity"
                    ),

                    close_path=drawing.get(
                        "closePath"
                    ),

                    items=item_list,
                )
            )

        return drawings

    def _extract_embedded_images(
        self,
        doc: fitz.Document,
        page: fitz.Page,
        page_number: int,
        assets_dir: Path,
    ) -> list[ImageAsset]:

        if not self.extract_embedded_images:
            return []

        assets = []

        seen_xrefs: set[int] = set()

        for (
            image_index,
            image,
        ) in enumerate(
            page.get_images(
                full=True
            ),
            start=1,
        ):

            xref = int(
                image[0]
            )

            if (
                xref in seen_xrefs
                or xref <= 0
            ):
                continue

            seen_xrefs.add(
                xref
            )

            try:
                extracted = (
                    doc.extract_image(
                        xref
                    )
                )

            except Exception:
                # Some PDF image references cannot
                # be directly extracted.
                #
                # The page rendering remains
                # available as fallback.
                continue

            extension = str(
                extracted.get(
                    "ext",
                    "bin",
                )
            )

            filename = (
                f"page_{page_number:03d}"
                f"_embedded_{image_index:03d}"
                f".{extension}"
            )

            output_path = (
                assets_dir / filename
            )

            output_path.write_bytes(
                extracted["image"]
            )

            bbox = None

            try:

                rects = page.get_image_rects(
                    xref
                )

                if rects:
                    bbox = self._bbox(
                        rects[0]
                    )

            except Exception:
                bbox = None

            assets.append(
                ImageAsset(
                    asset_id=(
                        f"p{page_number:03d}"
                        f"_img_{image_index:03d}"
                    ),

                    kind="embedded",

                    page_number=page_number,

                    path=str(
                        output_path
                    ),

                    width=(
                        int(
                            extracted["width"]
                        )
                        if extracted.get(
                            "width"
                        ) is not None
                        else None
                    ),

                    height=(
                        int(
                            extracted["height"]
                        )
                        if extracted.get(
                            "height"
                        ) is not None
                        else None
                    ),

                    extension=extension,

                    xref=xref,

                    bbox=bbox,

                    source="pdf_embedded_image",
                )
            )

        return assets

    def _render_page(
        self,
        page: fitz.Page,
        page_number: int,
        assets_dir: Path,
    ) -> ImageAsset:

        scale = (
            self.render_dpi / 72.0
        )

        matrix = fitz.Matrix(
            scale,
            scale,
        )

        pixmap = page.get_pixmap(
            matrix=matrix,
            alpha=False,
        )

        filename = (
            f"page_{page_number:03d}.png"
        )

        output_path = (
            assets_dir / filename
        )

        pixmap.save(
            str(output_path)
        )

        return ImageAsset(
            asset_id=(
                f"p{page_number:03d}"
                "_render"
            ),

            kind="page_render",

            page_number=page_number,

            path=str(
                output_path
            ),

            width=pixmap.width,

            height=pixmap.height,

            extension="png",

            source=(
                f"page_render_"
                f"{self.render_dpi}dpi"
            ),
        )

    def load(self) -> RawDocument:
        """
        Extract the complete Stage-1
        raw representation.
        """

        self.output_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        assets_dir = (
            self.output_dir / "assets"
        )

        assets_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        with fitz.open(
            self.pdf_path
        ) as doc:

            pages = []

            for (
                page_number,
                page,
            ) in enumerate(
                doc,
                start=1,
            ):

                (
                    elements,
                    word_count,
                    text_block_count,
                ) = self._extract_text_elements(
                    page
                )

                page_text = page.get_text(
                    "text",
                    sort=False,
                )

                text_char_count = len(
                    page_text
                )

                embedded_count = len(
                    page.get_images(
                        full=True
                    )
                )

                drawings = (
                    self._extract_drawings(
                        page
                    )
                )

                images = (
                    self._extract_embedded_images(
                        doc,
                        page,
                        page_number,
                        assets_dir,
                    )
                )

                if self.render_pages:

                    images.append(
                        self._render_page(
                            page,
                            page_number,
                            assets_dir,
                        )
                    )

                pages.append(
                    PageInfo(
                        page_number=page_number,

                        width=round(
                            float(
                                page.rect.width
                            ),
                            4,
                        ),

                        height=round(
                            float(
                                page.rect.height
                            ),
                            4,
                        ),

                        rotation=int(
                            page.rotation
                        ),

                        mediabox=self._bbox(
                            page.mediabox
                        ),

                        cropbox=self._bbox(
                            page.cropbox
                        ),

                        text_char_count=
                            text_char_count,

                        word_count=
                            word_count,

                        text_block_count=
                            text_block_count,

                        embedded_image_count=
                            embedded_count,

                        drawing_count=
                            len(drawings),

                        has_native_text=
                            text_char_count > 0,

                        needs_ocr_candidate=
                            text_char_count
                            < self.OCR_CHAR_THRESHOLD,

                        elements=
                            elements,

                        images=
                            images,

                        drawings=
                            drawings,
                    )
                )

            metadata = {
                str(key): value
                for key, value in (
                    doc.metadata or {}
                ).items()
                if value is not None
            }

            document_info = {
                "file_name":
                    self.pdf_path.name,

                "source_path":
                    str(self.pdf_path),

                "file_size_bytes":
                    self.pdf_path.stat().st_size,

                "sha256":
                    self._sha256(),

                "page_count":
                    doc.page_count,

                "pdf_metadata":
                    metadata,

                "extraction_engine":
                    "PyMuPDF",

                "extraction_engine_version":
                    fitz.VersionBind,

                "extracted_at_utc":
                    datetime.now(
                        timezone.utc
                    ).isoformat(),
            }

        return RawDocument(
            document=document_info,
            pages=pages,
        )


def save_raw_document(
    document: RawDocument,
    output_path: str | Path,
) -> Path:

    output = (
        Path(output_path)
        .expanduser()
        .resolve()
    )

    output.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    payload = document.model_dump(
        mode="json",
        exclude_none=False,
    )

    output.write_text(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    return output