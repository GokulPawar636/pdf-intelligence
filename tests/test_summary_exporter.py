from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app.reporting.summary_exporter import tolerance_limits, export_summary
from pipeline_main import run_pipeline


@pytest.mark.parametrize("nominal,tolerance,expected", [
    ("18", "±0.500", ("17.500", "18.500")),
    ("253.4", "+0.500/-1.500", ("251.9", "253.9")),
    ("91.3", "+4.000/0.000", ("91.3", "95.3")),
    ("18", "0.500", (None, None)),
    ("18", "�0.500", (None, None)),
    ("18", "+0.5/1.5", (None, None)),
])
def test_only_explicit_tolerances_produce_limits(nominal, tolerance, expected):
    assert tolerance_limits(Decimal(nominal), tolerance) == tuple(Decimal(v) if v is not None else None for v in expected)


def test_supplied_pdf_produces_one_sheet_with_verified_values(tmp_path):
    source = Path(__file__).resolve().parents[1] / "PDF1.pdf"
    output = run_pipeline(source, tmp_path, offline=True, report="summary")
    values = load_workbook(output, data_only=True)
    formulas = load_workbook(output, data_only=False)
    assert values.sheetnames == ["Summary Report", "Differences", "raw data"]
    sheet = values.active
    assert sheet["A20"].value == "Report overview"
    assert "12 characteristics | 12 readings" in sheet["C21"].value
    assert "Within tolerance: 12 | Out of tolerance: 0" in sheet["C22"].value
    assert "Review required" in sheet["C24"].value
    assert values["raw data"].max_row > 100
    assert values["raw data"]["A1"].value == "Source records and review notes (audit section)"
    # Independent transcription from the PDF, in page/table order.
    expected = [
        ("circle 3", "Z Distance", 18, 18.081, 0.5, "11"),
        ("circle 1", "Y Distance", 18, 17.913, 0.5, "8"),
        ("Ballooning No. 16", "Y Distance", 33, 32.888, 0.6, "9"),
        ("point 1", "Y Distance", 4.75, 4.585, 0.3, "7"),
        ("Ballooning No. 15", "Diameter", 30.5, 30.618, 1.2, "15"),
        ("Ballooning No. 16", "Diameter", 50, 49.941, 1.2, "16"),
        ("Ballooning No. 14", "Radius", 2.502, 2.689, 0.6, "14"),
        ("Ballooning No. 15", "X Distance", 78, 78.6, 0.8, "13"),
        ("point 4", "Z Distance", 28, 27.803, 0.6, "12"),
        ("Ballooning No. 18", "Radius", 2.5, 2.493, 0.6, "18"),
        ("point 2", "X Distance", 102.995, 102.704, 0.8, "19"),
        ("plane 2", "Y Distance", 33, 32.98, 0.6, "17"),
    ]
    for r, (name, control, nominal, measured, tolerance, balloon) in enumerate(expected, 7):
        assert sheet.cell(r, 1).value == name
        assert sheet.cell(r, 2).value == control
        assert sheet.cell(r, 3).value == nominal
        assert sheet.cell(r, 4).value == pytest.approx(nominal - tolerance)
        assert sheet.cell(r, 5).value == pytest.approx(nominal + tolerance)
        assert sheet.cell(r, 7).value == measured
        assert sheet.cell(r, 8).value == measured  # mean with one reading
        assert all(sheet.cell(r, c).value == "N/A" for c in (9, 10, 11))
        assert sheet.cell(r, 14).value == "OK"
        assert sheet.cell(r, 15).value == int(balloon)
        assert formulas.active.cell(r, 8).data_type == "f"
        assert formulas.active.cell(r, 7).comment is None
        assert "Record:" in sheet.cell(r, 22).value
    assert sheet["G6"].value == "8/25/2026"
    assert "8/25/2026" in sheet["A2"].value
    assert "Organization: UACPL" in sheet["A4"].value
    assert "Part name: HOLDING" in sheet["A4"].value
    assert "2025" not in str(sheet["G6"].value)
    differences = values["Differences"]
    assert differences["A4"].value == "Object"
    assert differences["F5"].value == pytest.approx(0.081)
    assert differences["J5"].value == "OK"
