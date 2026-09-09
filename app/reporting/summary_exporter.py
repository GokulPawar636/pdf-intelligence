"""Optional single-sheet measurement summary, based on the supplied sample.

Only fields present in the input are mapped. The general exporter remains
domain independent; no sample measurements or dates are used here.
"""
from __future__ import annotations

import json
import re
import statistics
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

from app.models.semantic import SemanticDocument
from app.reporting.storage import save_workbook_atomic
from app.reporting.excel_exporter import _append, extract_title_fields, fit_row_heights
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from app.models.structure import StructuredDocument


def number(value) -> Decimal | None:
    text = str(value).strip()
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", text):
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def tolerance_limits(nominal: Decimal | None, raw: str) -> tuple[Decimal | None, Decimal | None]:
    if nominal is None:
        return None, None
    text = re.sub(r"\s+", "", raw).replace("−", "-").replace("+/-", "±")
    if text.startswith("±") and (delta := number(text[1:])) is not None:
        return nominal - abs(delta), nominal + abs(delta)
    parts = text.split("/")
    if len(parts) == 2 and all(number(p) is not None for p in parts):
        # A nonzero unsigned offset is ambiguous; never assume its sign.
        if all(p.startswith(("+", "-")) or number(p) == 0 for p in parts):
            offsets = [number(p) for p in parts]
            return nominal + min(offsets), nominal + max(offsets)
    return None, None


@dataclass
class SummaryRow:
    object_name: str
    control: str
    nominal: Decimal | None
    lower: Decimal | None
    upper: Decimal | None
    tolerance: str
    measurements: list[Decimal]
    balloon: str = ""
    reference: str = ""
    page: str = ""
    source_result: str = ""
    evidence: str = ""
    warnings: list[str] = field(default_factory=list)
    unit: str = ""
    record_ids: list[str] = field(default_factory=list)
    explicit_limits: bool = False
    group_context: tuple = ()


def measurement_values(values: dict) -> list[Decimal]:
    readings = []
    for name, candidates in values.items():
        if not re.fullmatch(r"(?:meas|measured(?:_value)?|measurements?|actual(?:_value)?|readings?|sample|measurement_history)(?:_?\d+)?", name):
            continue
        for candidate in candidates:
            if isinstance(candidate, str) and candidate.strip().startswith("["):
                try:
                    candidate = json.loads(candidate)
                except (ValueError, TypeError):
                    pass
            items = candidate if isinstance(candidate, list) else [candidate]
            for item in items:
                if item in (None, ""):
                    continue
                value = number(item)
                if value is None or not value.is_finite() or abs(value) > Decimal("1e100"):
                    raise ValueError(f"Unreadable or unsupported measurement: {item!r}")
                if len(value.as_tuple().digits) > 15:
                    raise ValueError("Measurement exceeds Excel numeric precision (15 significant digits).")
                readings.append(value)
    return readings


def build_summary_rows(document: dict, structured: dict) -> list[SummaryRow]:
    tables = {t["table_id"]: (p, t) for p in structured.get("pages", []) for t in p.get("tables", [])}
    rows = []
    for record in document.get("records", []):
        values = {}
        for item in record.get("fields", []):
            raw = item["value"].get("raw_value")
            if raw is None:
                raw = item["value"].get("normalized_value")
            values.setdefault(re.sub(r"[^\w]+", "_", item["name"].casefold()).strip("_"), []).append(raw)

        def first(*names):
            return next((v for n in names for v in values.get(n, []) if v not in (None, "")), "")

        measurement_warning = ""
        try:
            measurements = measurement_values(values)
        except ValueError as exc:
            # Preserve valid readings but never present partial data as complete.
            measurements = []
            measurement_warning = str(exc)
            for name, candidates in values.items():
                for candidate in candidates:
                    try:
                        measurements.extend(measurement_values({name: [candidate]}))
                    except ValueError:
                        pass
        if not measurements:
            continue
        nominal = number(first("nom", "nominal", "nominal_value"))
        tolerance = str(first("tol", "tolerance"))
        lower, upper = tolerance_limits(nominal, tolerance)
        explicit_lower = number(first("lower_tolerance", "lower_limit", "lsl"))
        explicit_upper = number(first("upper_tolerance", "upper_limit", "usl"))
        lower = explicit_lower if explicit_lower is not None else lower
        upper = explicit_upper if explicit_upper is not None else upper
        context = ""
        table_id = (record.get("record_id") or "").rsplit("_r", 1)[0]
        if table_id in tables:
            page, table = tables[table_id]
            top = table["bbox"]["y0"]
            previous_bottom = max((t["bbox"]["y1"] for t in page["tables"]
                                   if t["bbox"]["y1"] < top), default=0)
            context = "\n".join(r["text"] for r in page["regions"]
                                if r["region_type"] == "text_region"
                                and previous_bottom <= r["bbox"]["y0"] < top
                                and top - r["bbox"]["y0"] <= 60)
        balloon_match = re.search(r"Ballooning\s+No\.?\s*(\d+)", context, re.I)
        feature_match = re.search(r"(?<!Ref )\bFeature:\s*(.+?)(?:,?\s+Ref Feature:|\n|$)", context, re.I)
        reference_match = re.search(r"Ref Feature:\s*([^\n]+)", context, re.I)
        balloon = str(first("balloon_number") or (balloon_match[1] if balloon_match else ""))
        object_name = str(first("object", "object_name", "feature") or
                          (feature_match[1].rstrip(", ") if feature_match else "") or
                          (f"Ballooning No. {balloon}" if balloon else "Not stated"))
        reference = str(first("reference_feature") or (reference_match[1].strip() if reference_match else ""))
        warnings = [measurement_warning] if measurement_warning else []
        if lower is None or upper is None:
            warnings.append("Limits unavailable: nominal or an explicit signed tolerance is missing.")
        if explicit_lower is not None or explicit_upper is not None:
            derived_lower, derived_upper = tolerance_limits(nominal, tolerance)
            if ((explicit_lower is not None and derived_lower is not None and explicit_lower != derived_lower)
                    or (explicit_upper is not None and derived_upper is not None and explicit_upper != derived_upper)):
                warnings.append("Explicit limits conflict with the tolerance text; verify the specification.")
        if lower is not None and upper is not None and lower > upper:
            warnings.append("Conflicting source limits; computed status omitted.")
            lower = upper = None
        source_result = str(first("test", "test_result", "result"))
        if lower is not None and upper is not None and source_result.casefold() in ("pass", "fail"):
            passed = lower <= measurements[-1] <= upper
            if passed != (source_result.casefold() == "pass"):
                warnings.append("Source test result disagrees with calculated tolerance check.")
        evidence = f"Record: {record.get('record_id')}\nContext: {context}\n" + "\n".join(
            f"{f['name']}: {f['value'].get('raw_value')}" for f in record["fields"])
        rows.append(SummaryRow(object_name, str(first("control", "axis", "characteristic")),
            nominal, lower, upper, tolerance, measurements, balloon, reference,
            ", ".join(map(str, record.get("source_pages", []))), source_result, evidence, warnings, str(first("unit", "units") or next((f["value"].get("unit") for f in record["fields"] if f["value"].get("unit")), ""))))
        rows[-1].record_ids = [record.get("record_id") or ""]
        rows[-1].explicit_limits = explicit_lower is not None or explicit_upper is not None
        rows[-1].group_context = tuple(str(first(name)) for name in
            ("part", "part_number", "component", "batch", "lot", "coordinate_system", "alignment"))
    # Only pool repeated rows when a source feature and control identify them.
    # Units, reference feature, balloon and specification must also agree.
    grouped = {}
    for index, row in enumerate(rows):
        key = (row.object_name, row.control, row.nominal, row.lower, row.upper,
               row.tolerance, row.balloon, row.reference, row.unit, row.group_context, row.explicit_limits)
        # Without reliable identity, keep observations separate.
        if row.object_name == "Not stated" or not row.control or not row.unit:
            key += (index,)
        if key not in grouped:
            grouped[key] = row
        else:
            existing = grouped[key]
            existing.measurements.extend(row.measurements)
            existing.page = ", ".join(dict.fromkeys((existing.page + ", " + row.page).split(", ")))
            existing.evidence += "\n" + row.evidence
            existing.warnings = list(dict.fromkeys(existing.warnings + row.warnings))
            existing.source_result = row.source_result
            existing.record_ids.extend(row.record_ids)
    return list(grouped.values())


def _cache_formulas(path: Path, caches: dict[str, float | str]) -> None:
    """Store preview values alongside formulas; Excel still recalculates on open."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", namespace)
    with ZipFile(path) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for info, data in entries:
            if info.filename == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                for cell in root.iter(f"{{{namespace}}}c"):
                    if cell.get("r") not in caches:
                        continue
                    cached = caches[cell.get("r")]
                    value = cell.find(f"{{{namespace}}}v")
                    if value is None:
                        value = ET.SubElement(cell, f"{{{namespace}}}v")
                    cell.set("t", "str" if isinstance(cached, str) else "n")
                    value.text = str(cached)
                data = ET.tostring(root, encoding="utf-8")
            archive.writestr(info, data)


def export_summary(semantic_path: str | Path, structured_path: str | Path, output_path: str | Path) -> Path:
    document = SemanticDocument.model_validate_json(Path(semantic_path).read_text(encoding="utf-8")).model_dump()
    structured = StructuredDocument.model_validate_json(Path(structured_path).read_text(encoding="utf-8")).model_dump()
    rows = build_summary_rows(document, structured)
    if not rows:
        raise ValueError("No measurement records found for a summary report. Use the dynamic report for this document.")
    source_name = document["document_metadata"].get("file_name", "Input PDF")
    source_text = "\n".join(r["text"] for p in structured["pages"] for r in p["regions"] if r["region_type"] == "text_region")
    date_match = re.search(r"\bDate:\s*([^\n]+)", source_text)
    date_label = date_match[1].strip() if date_match else "Date not stated"
    workbook = Workbook()
    workbook.calculation = CalcProperties(calcId=191029, fullCalcOnLoad=True, forceFullCalc=True, calcMode="auto")
    sheet = workbook.active
    sheet.title = "Summary Report"
    count = max(len(row.measurements) for row in rows)
    stat_start = 7 + count
    last_column = stat_start + 16
    end = get_column_letter(last_column)
    heading_end = get_column_letter(stat_start + 6)
    sheet.merge_cells(f"A1:{heading_end}1")
    sheet["A1"] = f"Summary Report - {Path(source_name).stem}"
    sheet["A1"].font = Font(name="Arial", size=16, bold=True, color="17365D")
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 30
    for row, text in [(2, f"Source: {source_name} | Inspection date: {date_label} | Measurements: {len(rows)}"),
                      (3, "Limits are nominal plus signed tolerance. Std Deviation = sample variation; UCL = Mean + 3 x Std Deviation; LCL = Mean - 3 x Std Deviation. N/A means fewer than 2 readings; variation cannot be estimated.")]:
        sheet.merge_cells(f"A{row}:{heading_end}{row}")
        sheet.cell(row, 1, text).alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[row].height = 28
    title_fields = extract_title_fields(document)
    title_text = " | ".join(f"{label}: {value}" for label, value in title_fields)
    if title_text:
        sheet.merge_cells(f"A4:{heading_end}4")
        sheet.cell(4, 1, title_text).alignment = Alignment(wrap_text=True, vertical="center")
        sheet.cell(4, 1).font = Font(bold=True, color="17365D")
        sheet.row_dimensions[4].height = 28
    for c, title in enumerate(["Object", "Control", "Nominal", "Lower Tolerance", "Upper Tolerance", "Tolerance"], 1):
        sheet.merge_cells(start_row=5, start_column=c, end_row=6, end_column=c)
        sheet.cell(5, c, title)
    for index in range(count):
        sheet.cell(5, 7 + index, "Measurement" if count == 1 else f"Measurement {index + 1}")
        sheet.cell(6, 7 + index, date_label if count == 1 else "Reading " + str(index + 1))
    sheet.merge_cells(start_row=5, start_column=stat_start, end_row=5, end_column=stat_start + 3)
    sheet.cell(5, stat_start, "Statistics")
    sheet.merge_cells(start_row=5, start_column=stat_start + 4, end_row=5, end_column=stat_start + 6)
    sheet.cell(5, stat_start + 4, "Spread Value")
    for offset, title in enumerate(["Mean", "Std Deviation (Variation)", "UCL (Upper Control Limit)", "LCL (Lower Control Limit)", "Min", "Max", "Remark"]):
        sheet.cell(6, stat_start + offset, title)
    for offset, title in enumerate(["Balloon No.", "Reference Feature", "Source Page", "Source Test", "Review Notes", "Readings", "Unit", "Evidence", "Lower Offset", "Upper Offset"], 7):
        column = stat_start + offset
        sheet.merge_cells(start_row=5, start_column=column, end_row=6, end_column=column)
        sheet.cell(5, column, title)
    caches = {}
    for r, row in enumerate(rows, 7):
        for c, value in enumerate([row.object_name, row.control, row.nominal, row.lower, row.upper, row.tolerance] + row.measurements, 1):
            if isinstance(value, str):
                value = ILLEGAL_CHARACTERS_RE.sub("", value)
            cell = sheet.cell(r, c, float(value) if isinstance(value, Decimal) else value)
            if isinstance(value, str):
                cell.data_type = "s"

        # Limits remain auditable and update if nominal changes.
        if row.nominal is not None and not row.explicit_limits:
            for column, limit in [(4, row.lower), (5, row.upper)]:
                if limit is not None:
                    delta = limit - row.nominal
                    cell = sheet.cell(r, column, f'=IF(C{r}="","",C{r}+{get_column_letter(stat_start + 15 if column == 4 else stat_start + 16)}{r})')
                    caches[cell.coordinate] = float(limit)

        measurement_end = get_column_letter(6 + count)
        span = f"G{r}:{measurement_end}{r}"
        vals = [float(v) for v in row.measurements]
        mean = statistics.mean(vals)
        deviation = statistics.stdev(vals) if len(vals) > 1 else None
        mean_cell = f"{get_column_letter(stat_start)}{r}"
        std_cell = f"{get_column_letter(stat_start + 1)}{r}"
        valid_limits = row.lower is not None and row.upper is not None
        status = "OK" if valid_limits and row.lower <= row.measurements[-1] <= row.upper else "Out of tolerance" if valid_limits else "Review"
        needs_review = bool(row.warnings)
        if needs_review:
            status = "Review"
        formulas = [
            (f'=IF(COUNT({span})=0,"",AVERAGE({span}))', mean),
            (f'=IF(COUNT({span})<2,"N/A",STDEV({span}))', deviation if deviation is not None else "N/A"),
            (f'=IF(COUNT({span})<2,"N/A",{mean_cell}+3*{std_cell})', mean + 3 * deviation if deviation is not None else "N/A"),
            (f'=IF(COUNT({span})<2,"N/A",{mean_cell}-3*{std_cell})', mean - 3 * deviation if deviation is not None else "N/A"),
            (f'=IF(COUNT({span})=0,"",MIN({span}))', min(vals)),
            (f'=IF(COUNT({span})=0,"",MAX({span}))', max(vals)),
            (f'=IF(OR({get_column_letter(stat_start + 11)}{r}<>"",COUNT({span})=0,D{r}="",E{r}=""),"Review",IF(AND(LOOKUP(9.99999999999999E+307,{span})>=D{r},LOOKUP(9.99999999999999E+307,{span})<=E{r}),"OK","Out of tolerance"))', status),
        ]
        for offset, (formula, cached) in enumerate(formulas):
            cell = sheet.cell(r, stat_start + offset, formula)
            caches[cell.coordinate] = cached
        for offset, value in enumerate([int(row.balloon) if row.balloon.isdigit() else row.balloon,
                                      row.reference, int(row.page) if row.page.isdigit() else row.page,
                                      row.source_result, "\n".join(row.warnings), len(vals), row.unit, row.evidence,
                                      float(row.lower - row.nominal) if row.lower is not None and row.nominal is not None else None,
                                      float(row.upper - row.nominal) if row.upper is not None and row.nominal is not None else None], 7):
            if isinstance(value, str):
                value = ILLEGAL_CHARACTERS_RE.sub("", value)
            cell = sheet.cell(r, stat_start + offset, value)
            if isinstance(value, str):
                cell.data_type = "s"
        sheet.row_dimensions[r].height = 30
        status_cell = sheet.cell(r, stat_start + 6)
        status_cell.fill = PatternFill("solid", fgColor="C6EFCE" if status == "OK" else "FFEB9C" if status == "Review" else "FFC7CE")
    border = Border(*( [Side(style="thin", color="A6A6A6")] * 4 ))
    for cells in sheet.iter_rows(min_row=5, max_row=6 + len(rows), max_col=last_column):
        for cell in cells:
            cell.border = border
            cell.font = Font(name="Arial", size=10, bold=cell.row <= 6 or cell.column <= 2)
            cell.alignment = Alignment(horizontal="left" if cell.column in (1, 2, last_column) else "center", vertical="center", wrap_text=True)
            if cell.row <= 6:
                cell.fill = PatternFill("solid", fgColor="DCE6F1")
            elif cell.row % 2 and cell.column != stat_start + 6:
                cell.fill = PatternFill("solid", fgColor="F5F7FA")
            if cell.row > 6 and 3 <= cell.column < stat_start + 6 and cell.column != 6:
                cell.number_format = "0.000"
    for row in (5, 6):
        sheet.row_dimensions[row].height = 42
    for column in range(1, last_column + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 16
    sheet.column_dimensions["A"].width = 25
    sheet.column_dimensions["B"].width = 19
    for offset in (1, 2, 3):
        sheet.column_dimensions[get_column_letter(stat_start + offset)].width = 19
    sheet.column_dimensions[get_column_letter(stat_start + 6)].width = 22
    fit_row_heights(sheet, 7, 6 + len(rows), stat_start + 6)
    sheet.column_dimensions[get_column_letter(stat_start + 8)].width = 23
    sheet.column_dimensions[end].width = 42
    remark_letter = get_column_letter(stat_start + 6)
    remark_range = f"{remark_letter}7:{remark_letter}{6 + len(rows)}"
    for value, color in [("OK", "C6EFCE"), ("Out of tolerance", "FFC7CE"), ("Review", "FFEB9C")]:
        sheet.conditional_formatting.add(remark_range, CellIsRule(operator="equal", formula=[f'"{value}"'], fill=PatternFill("solid", fgColor=color)))
    sheet.column_dimensions.group(get_column_letter(stat_start + 7), end, hidden=True)
    sheet.column_dimensions[get_column_letter(stat_start + 7)].collapsed = True
    sheet.freeze_panes = "C7"
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 80
    sheet.print_title_rows = "1:6"
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = f"A1:{heading_end}{6 + len(rows)}"
    # A single workbook/sheet keeps the source records available below the report.
    # They are outside the print area, not dropped by the summary projection.
    detail_row = sheet.max_row + 3
    sheet.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=6)
    sheet.cell(detail_row, 1, "Source records and review notes (audit section)").font = Font(bold=True, color="17365D")
    for note in document["warnings"] + document["global_notes"]:
        _append(sheet, ["Document note", note])
    _append(sheet, ["Record ID", "Record Type", "Field", "Raw Value", "Unit", "Source Page"])
    audit_header = sheet.max_row
    for record in document["records"]:
        for item in record["fields"]:
            value = item["value"]
            raw = value.get("raw_value")
            if raw is None:
                raw = str(value.get("normalized_value", ""))
            source = value.get("source") or {}
            for offset in range(0, max(1, len(raw)), 30000):
                _append(sheet, [record.get("record_id"), record["record_type"], item["name"],
                               raw[offset:offset + 30000], value.get("unit"), source.get("page_number")])
    for cell in sheet[audit_header][:6]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[audit_header].height = 32
    for cells in sheet.iter_rows(min_row=audit_header + 1, max_col=6):
        for cell in cells:
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
            if cell.row % 2:
                cell.fill = PatternFill("solid", fgColor="F5F7FA")
    fit_row_heights(sheet, audit_header + 1, sheet.max_row, 6)
    for cells in sheet:
        for cell in cells:
            if cell.data_type == "s" and isinstance(cell.value, str):
                cell.value = ILLEGAL_CHARACTERS_RE.sub("", cell.value)
                cell.data_type = "s"
    differences = workbook.create_sheet("Differences", 1)
    difference_end = get_column_letter(12)
    differences.merge_cells(f"A1:{difference_end}1")
    differences["A1"] = f"Differences - {Path(source_name).stem}"
    differences["A1"].font = Font(name="Arial", size=16, bold=True, color="17365D")
    differences["A1"].alignment = Alignment(horizontal="center", vertical="center")
    differences.merge_cells(f"A2:{difference_end}2")
    differences["A2"] = f"Source: {source_name} | Difference = Measurement - Nominal"
    differences["A2"].alignment = Alignment(wrap_text=True)
    header_row = 4 if title_text else 3
    if title_text:
        differences.merge_cells(f"A3:{difference_end}3")
        differences["A3"] = title_text
        differences["A3"].font = Font(bold=True, color="17365D")
        differences["A3"].alignment = Alignment(wrap_text=True)
    headers = ["Object", "Control", "Balloon No.", "Nominal", "Measurement", "Difference",
               "Tolerance", "Lower Limit", "Upper Limit", "Status", "Source Page", "Reference Feature"]
    _append(differences, headers)
    for row in rows:
        for measurement in row.measurements:
            status = "Review" if row.warnings or row.lower is None or row.upper is None else (
                "OK" if row.lower <= measurement <= row.upper else "Out of tolerance")
            _append(differences, [row.object_name, row.control,
                                  int(row.balloon) if row.balloon.isdigit() else row.balloon,
                                  float(row.nominal) if row.nominal is not None else None,
                                  float(measurement),
                                  float(measurement - row.nominal) if row.nominal is not None else None,
                                  row.tolerance,
                                  float(row.lower) if row.lower is not None else None,
                                  float(row.upper) if row.upper is not None else None,
                                  status, row.page, row.reference])
    differences.sheet_view.showGridLines = False
    differences.freeze_panes = f"A{header_row + 1}"
    for cell in differences[header_row]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row_cells in differences.iter_rows(min_row=header_row + 1):
        for cell in row_cells:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for column in range(1, len(headers) + 1):
        differences.column_dimensions[get_column_letter(column)].width = 16
    differences.column_dimensions["A"].width = 25
    differences.column_dimensions["B"].width = 19
    differences.column_dimensions["L"].width = 24
    differences.row_dimensions[1].height = 30
    differences.row_dimensions[2].height = 28
    if title_text:
        differences.row_dimensions[3].height = 42
    differences.row_dimensions[header_row].height = 36
    differences.column_dimensions["J"].width = 22
    differences.auto_filter.ref = f"A{header_row}:L{differences.max_row}"
    differences.sheet_view.zoomScale = 90
    differences.print_title_rows = f"1:{header_row}"
    differences.print_area = differences.dimensions
    differences.page_setup.orientation = "landscape"
    differences.page_setup.paperSize = differences.PAPERSIZE_A3
    differences.page_setup.fitToWidth = 1
    differences.page_setup.fitToHeight = 0
    differences.sheet_properties.pageSetUpPr.fitToPage = True
    for cells in differences.iter_rows(min_row=header_row + 1):
        for cell in cells:
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="left" if cell.column in (1, 2, 12) else "center")
            cell.fill = PatternFill("solid", fgColor="F5F7FA" if cell.row % 2 else "FFFFFF")
        status_cell = cells[9]
        status_cell.fill = PatternFill("solid", fgColor={"OK": "C6EFCE", "Review": "FFEB9C"}.get(status_cell.value, "FFC7CE"))
    fit_row_heights(differences, header_row + 1, differences.max_row)
    for row_index in range(header_row + 1, differences.max_row + 1):
        for column in range(4, 10):
            differences.cell(row_index, column).number_format = "0.000"
    return save_workbook_atomic(workbook, output_path, lambda path: _cache_formulas(path, caches))
