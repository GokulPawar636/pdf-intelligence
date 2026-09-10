"""Excel formulas for editable measurement reports (no macros required)."""
from copy import copy
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill


def reading_span(row, mean_column, sheet_prefix=""):
    # Anchors lie OUTSIDE the input area, so inserting before the first reading
    # or immediately before Mean expands the range after Excel adjusts references.
    return (f"INDEX({sheet_prefix}${row}:${row},1,COLUMN({sheet_prefix}$F{row})+1):"
            f"INDEX({sheet_prefix}${row}:${row},1,COLUMN({sheet_prefix}${letter(mean_column)}{row})-1)")


def offset_formula(row, upper=False):
    text = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(F{row}," ",""),"\u2212","-"),"+/-","\u00b1")'
    left = f'LEFT({text},FIND("/",{text})-1)'
    right = f'MID({text},FIND("/",{text})+1,255)'
    signed = lambda part: f'OR(LEFT({part},1)="+",LEFT({part},1)="-",VALUE({part})=0)'
    pair = f'IF(AND({signed(left)},{signed(right)}),{"MAX" if upper else "MIN"}(VALUE({left}),VALUE({right})),"")'
    return (f'=IFERROR(IF(LEFT({text},1)="\u00b1",'
            f'{"" if upper else "-"}ABS(VALUE(MID({text},2,255))),{pair}),"")')


def make_summary_live(workbook, rows, mean_column, caches, overview_start):
    sheet = workbook["Summary Report"]
    last_row = 6 + len(rows)
    count_letter = letter(mean_column + 12)
    status_letter = letter(mean_column + 6)
    notes_letter = letter(mean_column + 11)

    def put(row, column, formula, cached=None):
        cell = sheet.cell(row, column)
        previous = caches.get(cell.coordinate, cell.value)
        cell.value = formula
        caches[cell.coordinate] = previous if cached is None else cached

    for r, source in enumerate(rows, 7):
        span = reading_span(r, mean_column)
        invalid = f'COUNTA({span})<>COUNT({span})'
        mean, sd = f'{letter(mean_column)}{r}', f'{letter(mean_column + 1)}{r}'
        for offset, fn in ((0, "AVERAGE"), (4, "MIN"), (5, "MAX")):
            put(r, mean_column + offset, f'=IF({invalid},"Review",IF(COUNT({span})=0,"",{fn}({span})))')
        put(r, mean_column + 1, f'=IF({invalid},"Review",IF(COUNT({span})<2,"N/A",STDEV({span})))')
        for offset, sign in ((2, "+"), (3, "-")):
            put(r, mean_column + offset, f'=IF({invalid},"Review",IF(COUNT({span})<2,"N/A",{mean}{sign}3*{sd}))')
        put(r, mean_column + 12, f'=COUNT({span})', len(source.measurements))
        latest = f'LOOKUP(9.99999999999999E+307,{span})'
        put(r, mean_column + 6,
            f'=IFERROR(IF(OR({notes_letter}{r}<>"",{invalid},COUNT({span})=0,'
            f'NOT(ISNUMBER(D{r})),NOT(ISNUMBER(E{r})),D{r}>E{r}),"Review",'
            f'IF(AND({latest}>=D{r},{latest}<=E{r}),"OK","Out of tolerance")),"Review")')
        if not source.explicit_limits:
            for column, upper in ((4, False), (5, True)):
                helper = mean_column + (16 if upper else 15)
                offset = source.upper if upper else source.lower
                cache = float(offset - source.nominal) if offset is not None and source.nominal is not None else ""
                put(r, helper, offset_formula(r, upper), cache)
                helper_cell = f'{letter(helper)}{r}'
                # Unresolved source conflicts must not silently reappear as valid limits.
                conflict = any("conflict" in note.lower() or "unsupported" in note.lower() for note in source.warnings)
                guard = f'{notes_letter}{r}="",' if conflict else ''
                put(r, column, f'=IF(AND({guard}ISNUMBER(C{r}),ISNUMBER({helper_cell})),C{r}+{helper_cell},"")',
                    float(offset) if offset is not None else "")
        for c in [3, 6] + list(range(7, mean_column)):
            font = copy(sheet.cell(r, c).font)
            font.color = "0070C0"
            sheet.cell(r, c).font = font
        sheet.cell(r, 6).comment = Comment(
            "Editable tolerance: use +/-0.5, ±0.5, or +0.5/-0.2. Derived limits update automatically. "
            "If the PDF specifies absolute limits, edit Lower/Upper Tolerance directly instead.", "PDF Intelligence")

    counts = f'{count_letter}7:{count_letter}{last_row}'
    statuses = f'{status_letter}7:{status_letter}{last_row}'
    put(overview_start + 1, 3,
        f'=COUNTA(A7:A{last_row})&" characteristics | "&SUM({counts})&" readings | "&'
        f'COUNTIF({counts},">=2")&" characteristics have 2+ readings for variation calculations."')
    put(overview_start + 2, 3,
        f'="Within tolerance: "&COUNTIF({statuses},"OK")&" | Out of tolerance: "&'
        f'COUNTIF({statuses},"Out of tolerance")&" | Cannot determine / review: "&COUNTIF({statuses},"Review")')
    unit_range = f'{letter(mean_column + 13)}7:{letter(mean_column + 13)}{last_row}'
    with_units = sum(bool(row.unit) for row in rows)
    unit_text = f'Units recorded for: {with_units} characteristics | Missing: {len(rows) - with_units}. See Unit in the expandable detail columns.'
    put(overview_start + 3, 3,
        f'="Units recorded for: "&COUNTIF({unit_range},"<>")&" characteristics | Missing: "&'
        f'COUNTBLANK({unit_range})&". See Unit in the expandable detail columns."', unit_text)
    quality_cell = sheet.cell(overview_start + 4, 3)
    original_quality = quality_cell.value.replace('"', '""')
    put(overview_start + 4, 3,
        f'=IF(COUNTIF({statuses},"Review")>0,"Review required: inspect invalid or missing inputs and source Review Notes.","{original_quality}")')
    sheet.cell(overview_start + 5, 3,
        "Results check the latest reading. Differences includes original readings plus live latest-reading results. "
        "N/A means too few readings for SD and control limits. Control limits are not specification limits.")
    sheet.cell(overview_start + 6, 3,
        "Edit readings in blue cells. To add a reading/date column, insert an entire Excel column immediately before Mean "
        "(or inside the reading area), then enter numeric values. Keep formulas and the Mean column intact. "
        "Review Notes refer to the original PDF; clear a note only after verifying its issue.")
    sheet.row_dimensions[overview_start + 6].height = 54
    validation = DataValidation(type="decimal", operator="between", formula1="-1E100", formula2="1E100", allow_blank=True)
    validation.errorTitle = "Enter a numeric measurement"
    validation.error = "Use numbers only. Leave missing measurements blank."
    validation.showErrorMessage = True
    validation.errorStyle = "stop"
    sheet.add_data_validation(validation)
    validation.add(f'G7:{letter(mean_column - 1)}{last_row}')

    # Existing per-reading differences remain linked to their source reading.
    # New reading columns contribute to the summary and overview immediately.
    differences = workbook["Differences"]
    header = next(cell.row for cell in differences['A'] if cell.value == 'Object')
    target_row = header + 1
    prefix = "'Summary Report'!"
    for source_row, source in enumerate(rows, 7):
        for reading_index, _ in enumerate(source.measurements):
            for dest, origin in ((1, 1), (2, 2), (3, mean_column + 7), (4, 3),
                                 (5, 7 + reading_index), (7, 6), (8, 4), (9, 5),
                                 (11, mean_column + 9), (12, mean_column + 8)):
                cell = differences.cell(target_row, dest)
                caches[f'2!{cell.coordinate}'] = cell.value if cell.value is not None else ""
                ref = f'{prefix}{letter(origin)}{source_row}'
                cell.value = f'=IF({ref}="","",{ref})'
            cell = differences.cell(target_row, 6)
            caches[f'2!{cell.coordinate}'] = cell.value if cell.value is not None else ""
            cell.value = f'=IF(AND(ISNUMBER(D{target_row}),ISNUMBER(E{target_row})),E{target_row}-D{target_row},"")'
            cell = differences.cell(target_row, 10)
            caches[f'2!{cell.coordinate}'] = cell.value
            cell.value = (f'=IFERROR(IF(OR({prefix}{notes_letter}{source_row}<>"",'
                          f'NOT(ISNUMBER(E{target_row})),NOT(ISNUMBER(H{target_row})),NOT(ISNUMBER(I{target_row})),'
                          f'H{target_row}>I{target_row}),"Review",IF(AND(E{target_row}>=H{target_row},'
                          f'E{target_row}<=I{target_row}),"OK","Out of tolerance")),"Review")')
            target_row += 1
    latest_title = target_row + 1
    differences.merge_cells(start_row=latest_title, start_column=1, end_row=latest_title, end_column=12)
    differences.cell(latest_title, 1, "Latest readings - includes future reading columns")
    differences.cell(latest_title, 1)._style = copy(differences['A1']._style)
    differences.row_dimensions[latest_title].height = 30
    latest_header = latest_title + 1
    for c in range(1, 13):
        differences.cell(latest_header, c, differences.cell(header, c).value)
        differences.cell(latest_header, c)._style = copy(differences.cell(header, c)._style)
    differences.cell(latest_header, 5, "Latest Measurement")
    differences.row_dimensions[latest_header].height = 36
    for r, source in enumerate(rows, 7):
        dest_row = latest_header + r - 6
        def latest_put(column, formula, cached):
            cell = differences.cell(dest_row, column, formula)
            cell._style = copy(differences.cell(header + 1, column)._style)
            caches[f'2!{cell.coordinate}'] = cached if cached is not None else ""
        for dest, origin in ((1, 1), (2, 2), (3, mean_column + 7), (4, 3), (7, 6),
                             (8, 4), (9, 5), (10, mean_column + 6), (11, mean_column + 9), (12, mean_column + 8)):
            ref = f'{prefix}{letter(origin)}{r}'
            cell = sheet.cell(r, origin)
            latest_put(dest, f'=IF({ref}="","",{ref})', caches.get(cell.coordinate, cell.value))
        span = reading_span(r, mean_column, prefix)
        latest_put(5, f'=IF(COUNTA({span})<>COUNT({span}),"Review",IF(COUNT({span})=0,"",LOOKUP(9.99999999999999E+307,{span})))',
                   float(source.measurements[-1]))
        latest_put(6, f'=IF(AND(ISNUMBER(D{dest_row}),ISNUMBER(E{dest_row})),E{dest_row}-D{dest_row},"")',
                   float(source.measurements[-1] - source.nominal) if source.nominal is not None else "")
        differences.row_dimensions[dest_row].height = 30
    target_row = differences.max_row + 1
    differences.print_area = differences.dimensions
    # Compact reports fit on one page; larger reports retain readable pagination.
    differences.page_setup.fitToHeight = 1 if differences.max_row <= 40 else 0
    for value, color in (("OK", "C6EFCE"), ("Review", "FFEB9C"), ("Out of tolerance", "FFC7CE")):
        differences.conditional_formatting.add(f'J{header + 1}:J{target_row - 1}',
            CellIsRule(operator="equal", formula=[f'"{value}"'], fill=PatternFill("solid", fgColor=color)))
    differences["A2"] = "Top: original PDF reading positions, linked to Summary Report. Below: latest readings, including newly inserted reading columns."
