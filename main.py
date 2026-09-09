from __future__ import annotations

import argparse
from pathlib import Path

from app.ingestion.pdf_loader import PDFLoader, save_raw_document


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Stage 1: convert a PDF into a "
            "layout-preserving raw JSON representation."
        )
    )

    parser.add_argument(
        "pdf",
        help="Path to the input PDF",
    )

    parser.add_argument(
        "--output-dir",
        default="data/intermediate",
        help=(
            "Directory for JSON and extracted assets "
            "(default: data/intermediate)"
        ),
    )

    parser.add_argument(
        "--json-name",
        default="raw_document.json",
        help="Name of the generated JSON file",
    )

    parser.add_argument(
        "--no-render",
        action="store_true",
        help="Do not render each PDF page to PNG",
    )

    parser.add_argument(
        "--no-embedded-images",
        action="store_true",
        help="Do not extract embedded raster images",
    )

    parser.add_argument(
        "--no-drawings",
        action="store_true",
        help="Do not preserve vector drawing objects",
    )

    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="Page rendering DPI (default: 150)",
    )

    return parser


def main() -> None:
    args = build_parser().parse_args()

    output_dir = Path(args.output_dir)

    loader = PDFLoader(
        args.pdf,
        output_dir,
        render_pages=not args.no_render,
        extract_embedded_images=not args.no_embedded_images,
        include_drawings=not args.no_drawings,
        render_dpi=args.dpi,
    )

    document = loader.load()

    json_path = save_raw_document(
        document,
        output_dir / args.json_name,
    )

    print("Stage 1 extraction completed successfully.")
    print(f"Input PDF : {loader.pdf_path}")
    print(f"Pages     : {document.document.page_count}")
    print(f"JSON      : {json_path}")
    print()

    for page in document.pages:
        print(
            f"Page {page.page_number}: "
            f"chars={page.text_char_count}, "
            f"words={page.word_count}, "
            f"text_blocks={page.text_block_count}, "
            f"embedded_images={page.embedded_image_count}, "
            f"drawings={page.drawing_count}, "
            f"ocr_candidate={page.needs_ocr_candidate}"
        )


if __name__ == "__main__":
    main()