from __future__ import annotations

import argparse

from app.reporting.excel_exporter import export_semantic_to_excel


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Stage 4: export document-driven sheets and columns to Excel."
    )
    parser.add_argument("semantic_json", help="Path to Stage-3 semantic_document.json")
    parser.add_argument(
        "--output",
        default="data/output/document_results.xlsx",
        help="Output Excel workbook path",
    )
    args = parser.parse_args()

    output = export_semantic_to_excel(args.semantic_json, args.output)
    print("Stage 4 Excel export completed successfully.")
    print(f"Output: {output}")


if __name__ == "__main__":
    main()
