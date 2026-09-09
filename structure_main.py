from __future__ import annotations

import argparse

from app.structure.analyzer import (
    analyze_file,
)


def main() -> None:

    parser = argparse.ArgumentParser(
        description=(
            "Stage 2: dynamically detect "
            "document structure."
        )
    )

    parser.add_argument(
        "raw_json",
        help=(
            "Path to Stage-1 "
            "raw_document.json"
        ),
    )

    parser.add_argument(
        "--output",
        default=(
            "data/intermediate/"
            "structured_document.json"
        ),
        help=(
            "Output structured JSON path"
        ),
    )

    args = parser.parse_args()

    output = analyze_file(
        args.raw_json,
        args.output,
    )

    print(
        "Stage 2 structure analysis "
        "completed successfully."
    )

    print(
        f"Output: {output}"
    )


if __name__ == "__main__":
    main()