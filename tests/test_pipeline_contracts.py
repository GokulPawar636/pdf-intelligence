import json

import pytest
from pydantic import ValidationError

from app.semantic.extractor import SemanticExtractor, extract_file
from app.structure.analyzer import analyze_file


def _extractor(**kwargs):
    """Create an extractor without making a network request."""
    return SemanticExtractor(api_key="test-key", **kwargs)


def test_semantic_extractor_rejects_non_positive_chunk_size():
    with pytest.raises(ValueError, match="max_input_chars"):
        _extractor(max_input_chars=0)


def test_table_candidate_regions_are_not_sent_twice():
    extractor = _extractor()
    chunks = extractor._make_chunks(
        {
            "regions": [
                {"region_type": "text_region", "text": "title"},
                {
                    "region_type": "table_candidate",
                    "text": "duplicate table text",
                },
            ],
            "tables": [
                {
                    "table_id": "p001_t001",
                    "cells": [
                        {"row": 0, "column": 0, "text": "cell"}
                    ],
                }
            ],
        }
    )

    assert len(chunks) == 1
    assert [r["text"] for r in chunks[0]["regions"]] == ["title"]
    assert chunks[0]["tables"] == []


def test_chunker_enforces_the_rendered_prompt_budget():
    extractor = _extractor(max_input_chars=8000)
    chunks = extractor._make_chunks(
        {
            "page_number": 1,
            "regions": [
                {
                    "region_type": "text_region",
                    "text": "measurement " * 1800,
                    "bbox": {"x0": 0, "y0": 0, "x1": 1, "y1": 1},
                }
            ],
            "tables": [],
        }
    )

    assert len(chunks) > 1
    assert all(
        extractor._prompt_size(1, chunk) <= extractor.max_input_chars
        for chunk in chunks
    )


@pytest.mark.parametrize(
    ("header", "expected"),
    [("2.5", 2.5), ("1m2.5s", 62.5), ("3s", 3.0)],
)
def test_retry_delay_parser(header, expected):
    assert SemanticExtractor._parse_retry_delay(header) == expected


def test_offline_mode_does_not_require_a_groq_key():
    extractor = SemanticExtractor(api_key=None, offline=True)
    result = extractor.extract({"pages": []})

    assert result.records == []
    assert "offline mode" in result.global_notes[0].lower()


def test_generic_table_becomes_dynamic_table_row_records():
    extractor = _extractor()
    records = extractor._extract_table_records(
        {
            "pages": [{
                "page_number": 1,
                "tables": [{
                    "table_id": "p001_t001",
                    "cells": [
                        {"row": 0, "column": 0, "text": "Name"},
                        {"row": 0, "column": 1, "text": "Amount"},
                        {"row": 1, "column": 0, "text": "Widget"},
                        {"row": 1, "column": 1, "text": "42.5"},
                    ],
                }],
                "regions": [],
            }]
        }
    )

    assert records[0].record_type == "table_row"
    fields = {field.name: field.value.normalized_value for field in records[0].fields}
    assert fields == {"name": "Widget", "amount": 42.5}


def test_empty_model_records_are_discarded():
    extractor = _extractor()
    assert extractor._normalize_model_output(
        {"records": [{"record_type": "header", "fields": []}]}
    ) == {"records": []}


def test_semantic_document_uses_source_hash_as_a_stable_id():
    extractor = _extractor()
    result = extractor._merge_results(
        {
            "schema_name": "pdf-intelligence.structured-document",
            "schema_version": "2.0.0",
            "document": {"sha256": "abc123"},
        },
        [],
    )

    assert result.document_id == "abc123"
    assert result.schema_version == "3.1.0"


def test_stage_two_validates_stage_one_contract(tmp_path):
    invalid = tmp_path / "invalid-raw.json"
    invalid.write_text(json.dumps({"pages": []}), encoding="utf-8")

    with pytest.raises(ValidationError):
        analyze_file(invalid, tmp_path / "structured.json")


def test_stage_three_validates_stage_two_contract(tmp_path):
    invalid = tmp_path / "invalid-structured.json"
    invalid.write_text(json.dumps({"pages": []}), encoding="utf-8")

    with pytest.raises(ValidationError):
        extract_file(invalid, tmp_path / "semantic.json")
