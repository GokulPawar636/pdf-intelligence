"""Regression cases for silent sign changes, conflicting specs, and hidden uncertainty."""
import json
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app.semantic.extractor import SemanticExtractor
from app.reporting.summary_exporter import build_summary_rows, export_summary
from pipeline_main import run_pipeline


@pytest.mark.parametrize("source,claimed", [
    ("-10", "10"), ("+10", "10"), ("\u221210", "10"),
    ("1,000", "1"), ("10.25", "10"), ("110", "10"),
])
def test_ai_cannot_drop_signs_or_take_numeric_fragments(source, claimed):
    extractor = SemanticExtractor(offline=True)
    record = extractor._convert_record({"record_type": "details", "fields": [
        {"name": "label", "raw_value": "Sample", "evidence_ids": ["E0001"]},
        {"name": "reading", "raw_value": claimed, "evidence_ids": ["E0001"]},
    ]}, {"regions": [{"text": f"Sample reading: {source}", "region_id": "r1"}]}, 1, 1, 0)
    assert [field.name for field in record.fields] == ["label"]
    assert record.warnings


def test_exact_signed_value_survives_evidence_validation():
    record = SemanticExtractor(offline=True)._convert_record({"record_type": "details", "fields": [
        {"name": "reading", "raw_value": "-10", "evidence_ids": ["E0001"]},
    ]}, {"regions": [{"text": "Reading: -10", "region_id": "r1"}]}, 1, 1, 0)
    assert record.fields[0].value.normalized_value == -10


@pytest.mark.parametrize("extra", [
    [("nominal_value", "20")], [("lower_limit", "bad")],
    [("upper_limit", "1234567890123456")], [("tol", "+/-2")],
])
def test_bad_specs_cannot_produce_pass_status(extra):
    pairs = [("object", "Shaft"), ("control", "Diameter"), ("nominal", "10"),
             ("tolerance", "+/-0.5"), ("reading", "10")] + extra
    record = {"fields": [{"name": name, "value": {"raw_value": raw}} for name, raw in pairs]}
    row = build_summary_rows({"records": [record]}, {})[0]
    assert row.warnings
    assert row.lower is None and row.upper is None


def test_offline_uncertainty_reaches_manifest_and_strict_gate(tmp_path):
    source = Path(__file__).resolve().parents[1] / "PDF1.pdf"
    output = run_pipeline(source, tmp_path / "normal", offline=True)
    manifest = json.loads(output.with_name("manifest.json").read_text())
    assert manifest["status"] == "needs_review"
    assert manifest["quality"]["measurement_rows"] == 12
    assert manifest["quality"]["measurement_readings"] == 12
    assert manifest["quality"]["accuracy_verified"] is False
    assert any("headers" in warning for warning in manifest["warnings"])
    with pytest.raises(ValueError, match="Strict quality check failed"):
        run_pipeline(source, tmp_path / "strict", offline=True, strict=True)
    assert not list((tmp_path / "strict").rglob("*.xlsx"))


def test_differences_checks_each_reading_independently(tmp_path):
    pairs = [("object", "Shaft"), ("control", "Diameter"), ("nominal", "10"),
             ("lower_limit", "9.5"), ("upper_limit", "10.5"),
             ("reading_1", "9"), ("reading_2", "10")]
    semantic = {"source_schema": "test", "source_schema_version": "1", "document_id": "test",
                "extraction_model": "local", "records": [{"record_type": "measurement",
                "fields": [{"name": name, "value": {"raw_value": raw}} for name, raw in pairs]}]}
    structured = {"source_schema": "test", "source_schema_version": "1", "document": {}, "pages": []}
    source, layout = tmp_path / "semantic.json", tmp_path / "layout.json"
    source.write_text(json.dumps(semantic), encoding="utf-8")
    layout.write_text(json.dumps(structured), encoding="utf-8")
    workbook = load_workbook(export_summary(source, layout, tmp_path / "report.xlsx"), data_only=True)
    try:
        assert workbook["Differences"]["J4"].value == "Out of tolerance"
        assert workbook["Differences"]["J5"].value == "OK"
    finally:
        workbook.close()
