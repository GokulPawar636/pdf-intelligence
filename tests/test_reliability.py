import json
from pathlib import Path

import fitz
from openpyxl import Workbook
import pytest

from app.reporting import storage
from app.reporting.summary_exporter import build_summary_rows, measurement_values
import pipeline_main
from app.semantic.extractor import SemanticExtractor


def make_pdf(path, text="Project: Aurora"):
    pdf = fitz.open()
    page = pdf.new_page()
    if text:
        page.insert_text((50, 70), text)
    pdf.save(path)
    pdf.close()
    return path


def test_failed_publish_preserves_existing_file_and_cleans_temporary(tmp_path, monkeypatch):
    target = tmp_path / "output.xlsx"
    target.write_bytes(b"previous output")
    def locked(*args):
        raise PermissionError("locked")
    monkeypatch.setattr(storage.os, "replace", locked)
    with pytest.raises(PermissionError, match="previous file was preserved"):
        storage.save_workbook_atomic(Workbook(), target)
    assert target.read_bytes() == b"previous output"
    assert not list(tmp_path.glob(".building-*"))


def test_invalid_workbook_never_published(tmp_path):
    workbook = Workbook()
    workbook.active["A1"] = "#DIV/0!"
    with pytest.raises(ValueError, match="Excel error"):
        storage.save_workbook_atomic(workbook, tmp_path / "output.xlsx")
    assert not (tmp_path / "output.xlsx").exists()


def test_runs_are_isolated_and_manifest_matches_output(tmp_path):
    source = make_pdf(tmp_path / "input.pdf")
    first = pipeline_main.run_pipeline(source, tmp_path / "out", offline=True)
    previous_bytes = first.read_bytes()
    second = pipeline_main.run_pipeline(source, tmp_path / "out", offline=True)
    assert first != second
    assert first.read_bytes() == previous_bytes
    manifest = json.loads((second.parent / "manifest.json").read_text())
    assert manifest["status"] == "complete"
    assert manifest["selected_report"] == "dynamic"
    assert manifest["output"] == str(second)
    assert len(manifest["output_sha256"]) == 64
    assert not list(second.parent.glob(".building-*"))


def test_strict_mode_rejects_empty_extraction_and_records_failure(tmp_path):
    source = make_pdf(tmp_path / "scan.pdf", "")
    with pytest.raises(ValueError, match="Strict quality check failed"):
        pipeline_main.run_pipeline(source, tmp_path / "out", offline=True, strict=True)
    manifest = json.loads(next((tmp_path / "out").rglob("manifest.json")).read_text())
    assert manifest["status"] == "failed"
    assert manifest["warnings"]
    assert not list((tmp_path / "out").rglob("*.xlsx"))


def test_failed_stage_has_no_successful_output(tmp_path, monkeypatch):
    source = make_pdf(tmp_path / "input.pdf")
    def broken(*args, **kwargs):
        raise RuntimeError("detector failure")
    monkeypatch.setattr(pipeline_main, "analyze_file", broken)
    with pytest.raises(RuntimeError, match="detector failure"):
        pipeline_main.run_pipeline(source, tmp_path / "out", offline=True)
    manifest = json.loads(next((tmp_path / "out").rglob("manifest.json")).read_text())
    assert manifest["status"] == "failed" and manifest["stage"] == "structure"
    assert not list((tmp_path / "out").rglob("*.xlsx"))


def test_encrypted_and_oversized_inputs_rejected_before_run(tmp_path):
    source = tmp_path / "encrypted.pdf"
    pdf = fitz.open()
    pdf.new_page()
    pdf.save(source, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="secret", owner_pw="owner")
    pdf.close()
    with pytest.raises(ValueError, match="Password-protected"):
        pipeline_main.run_pipeline(source, tmp_path / "out", offline=True)
    large = tmp_path / "large.pdf"
    large.write_bytes(b"0" * (1024 * 1024 + 1))
    with pytest.raises(ValueError, match="1 MB"):
        pipeline_main.run_pipeline(large, tmp_path / "out", offline=True, max_mb=1)
    assert not (tmp_path / "out").exists()


@pytest.mark.parametrize("raw", ["not readable", "10,123", "NaN", "[1, \"broken\"]", "1234567890123456"])
def test_bad_measurements_do_not_silently_pass(raw):
    with pytest.raises(ValueError):
        measurement_values({"measurement": [raw]})


def test_partial_measurements_and_conflicting_limits_are_flagged():
    values = [("object", "A"), ("control", "Diameter"), ("nominal", "10"),
              ("tolerance", "+/-0.5"), ("lower_limit", "8"),
              ("measurement_1", "10"), ("measurement_2", "unreadable")]
    row = build_summary_rows({"records": [{"fields": [
        {"name": name, "value": {"raw_value": raw}} for name, raw in values]}]}, {})[0]
    assert row.explicit_limits
    assert any("Unreadable" in w for w in row.warnings)
    assert any("conflict" in w for w in row.warnings)
    assert list(map(float, row.measurements)) == [10]


def test_batch_continues_after_one_bad_input(tmp_path, monkeypatch):
    source = make_pdf(tmp_path / "valid.pdf")
    monkeypatch.setattr("sys.argv", ["pipeline_main.py", str(tmp_path / "missing.pdf"),
                                     str(source), "--offline", "--output-dir", str(tmp_path / "out")])
    assert pipeline_main.main() == 1
    assert len(list((tmp_path / "out").rglob("*.xlsx"))) == 1


def test_ai_values_must_match_evidence_and_cannot_change_identifiers():
    extractor = SemanticExtractor(offline=True)
    chunk = {"regions": [{"text": "Account ID: 00123. Amount: 18.081", "region_id": "r1"}]}
    record = extractor._convert_record({"record_type": "details", "fields": [
        {"name": "account_id", "raw_value": "00123", "normalized_value": 123, "evidence_ids": ["E0001"]},
        {"name": "amount", "raw_value": "99", "normalized_value": 99, "evidence_ids": ["E0001"]},
        {"name": "missing", "raw_value": "invented", "evidence_ids": []},
    ]}, chunk, 1, 1, 0)
    assert record.fields[0].name == "account_id"
    assert record.fields[0].value.normalized_value == "00123"
    assert all(f.name != "amount" and f.name != "missing" for f in record.fields)
    assert record.warnings


def test_ai_empty_result_keeps_original_text(monkeypatch):
    extractor = SemanticExtractor(api_key="test-key")
    monkeypatch.setattr(extractor, "_call_groq", lambda **kwargs: {"records": []})
    result = extractor.extract({"pages": [{"page_number": 1, "regions": [
        {"region_type": "text_region", "region_id": "r1", "text": "Project: Aurora"}], "tables": []}]})
    assert any(f.value.raw_value == "Aurora" for r in result.records for f in r.fields)
