from __future__ import annotations

import argparse

from app.semantic.extractor import extract_file


def main() -> None:

    parser = argparse.ArgumentParser(
        description=(
            "Stage 3: semantic understanding "
            "using Groq."
        )
    )

    parser.add_argument(
        "structured_json",
        help=(
            "Path to Stage-2 "
            "structured_document.json"
        ),
    )

    parser.add_argument(
        "--output",
        default=(
            "data/intermediate/"
            "semantic_document.json"
        ),
        help=(
            "Output semantic JSON path"
        ),
    )

    parser.add_argument(
        "--model",
        default="qwen/qwen3.8-27b",
        help="Groq model name",
    )

    parser.add_argument(
        "--max-input-chars",
        type=int,
        default=10000,
        help=(
            "Maximum characters in the complete Groq request "
            "(instructions and evidence included)."
        ),
    )

    parser.add_argument(
        "--max-output-tokens",
        type=int,
        default=500,
        help=(
            "Maximum Groq completion tokens."
        ),
    )

    parser.add_argument(
        "--offline",
        action="store_true",
        help=(
            "Skip Groq and extract tables, labeled values, and text "
            "records extracted from PDF evidence."
        ),
    )

    args = parser.parse_args()

    output = extract_file(
        input_path=args.structured_json,
        output_path=args.output,
        model=args.model,
        max_input_chars=args.max_input_chars,
        max_output_tokens=args.max_output_tokens,
        offline=args.offline,
    )

    print(
        "Stage 3 semantic extraction "
        "completed successfully."
    )

    print(
        f"Output: {output}"
    )


if __name__ == "__main__":
    main()
