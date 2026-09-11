from decimal import Decimal
import fitz
import pytest
from app.reporting.feature_tables import read_feature_tables
from app.batch import source_details


def feature_pdf(path, measured="8.009", result="Pass"):
    doc = fitz.open()
    page = doc.new_page()
    for y, text in [(30, "Part #: P123"), (45, "Serial #: 02"),
                    (60, "Length Units Millimeters"), (90, "Ballooning No.16,17")]:
        page.insert_text((30, y), text, fontsize=9)
    positions = [30, 200, 270, 355, 415, 460]
    for y, values in [(110, ["Control", "Nom", "Meas", "Tol", "Dev", "Test"]),
                      (130, ["8.000", "8.000", measured, "+0.035/-0.020", "0.009", result]),
                      (150, ["0.300 A D C", "", "0.239", "0.300", "0.239", result])]:
        for x, value in zip(positions, values):
            # Match the numeric columns' right-aligned layout.
            if value:
                if x == 355:
                    x -= max(0, len(value)-3)*4
                page.insert_text((x, y), value, fontsize=8)
    doc.save(path)
    doc.close()
    return path


def test_native_feature_table_keeps_blank_nominal_and_multi_balloon(tmp_path):
    path = feature_pdf(tmp_path / "features.pdf")
    rows = read_feature_tables(path)
    assert len(rows) == 2
    assert rows[0].balloon == rows[1].balloon == "16,17"
    assert rows[0].measurements == [Decimal("8.009")]
    assert rows[0].lower == Decimal("7.980")
    assert rows[0].upper == Decimal("8.035")
    assert rows[1].nominal is None
    assert rows[1].measurements == [Decimal("0.239")]
    assert rows[1].lower is None and rows[1].upper is None
    assert "symbol review" in rows[1].control
    assert source_details(path)["part"] == "P123"
    assert source_details(path)["serial"] == "02"


def test_unreadable_measurement_does_not_silently_disappear(tmp_path):
    path = feature_pdf(tmp_path / "bad.pdf", "???")
    with pytest.raises(ValueError, match="Unreadable feature-table measurement"):
        read_feature_tables(path)


def test_unrelated_document_is_not_a_feature_table(tmp_path):
    path = tmp_path / "note.pdf"
    doc = fitz.open()
    doc.new_page().insert_text((30, 30), "Measurements need review; Pass")
    doc.save(path)
    doc.close()
    assert read_feature_tables(path) == []


def test_measurement_rows_with_blank_test_are_preserved(tmp_path):
    rows = read_feature_tables(feature_pdf(tmp_path/'blank-test.pdf', result=''))
    assert len(rows) == 2
    assert rows[0].measurements == [Decimal('8.009')]
    assert rows[1].measurements == [Decimal('0.239')]
