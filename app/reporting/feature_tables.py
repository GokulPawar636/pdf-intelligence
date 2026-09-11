"""Read native feature tables by their printed column positions.

Only explicit Control/Nom/Meas/Tol/Dev/Test tables are recognized. Drawing
annotations are excluded; missing graphic symbols are never named by guessing.
"""
import re
import fitz
from .summary_exporter import SummaryRow, number, tolerance_limits


def read_feature_tables(path):
    rows = []
    active_alignment = active_coordinate = ''
    with fitz.open(path) as document:
        for page_index, page in enumerate(document):
            words = page.get_text("words")
            page_text = page.get_text(sort=True)
            alignments = re.findall(r'Data Alignments\s+([^\n]+)', page_text)
            coordinates = re.findall(r'Coordinate Systems\s+([^\n]+)', page_text)
            markers = [w for w in words if w[4] == "Control"]
            for marker in markers:
                header = sorted((w for w in words if abs(w[1] - marker[1]) < 2), key=lambda w: w[0])
                labels = [w[4] for w in header]
                if labels[:6] != ["Control", "Nom", "Meas", "Tol", "Dev", "Test"]:
                    continue
                # Feature-table headers only, never small drawing callouts.
                context_words = sorted((w for w in words if marker[1]-28 <= w[1] < marker[1]-2), key=lambda w: w[0])
                context = " ".join(w[4] for w in context_words)
                balloon = re.search(r"Ballooning\s+No\.?\s*(\d+[A-Za-z]?(?:,\d+[A-Za-z]?)*)", context, re.I)
                if not balloon:
                    continue
                feature = re.search(r"(?<!Ref )Feature:\s*(.*?)(?:,?\s+Ref Feature:|$)", context)
                reference = re.search(r"Ref Feature:\s*(.*)", context)
                # Numeric values are right aligned under the numeric headers.
                centers = [(w[0]+w[2])/2 for w in header[:6]]
                boundaries = [(centers[i]+centers[i+1])/2 for i in range(5)]
                end = min((w[1] for w in words if w[1] > marker[1]+3 and
                           (w[4] == "Ballooning" or w[4] == "Control")), default=page.rect.height-25)
                candidates = [w for w in words if marker[3] < w[1] < end]
                # Anchor rows on the explicit source result, retaining blank Nom cells.
                results = sorted((w for w in candidates if w[4].lower() in ("pass", "fail")
                                  and abs(w[0]-header[5][0]) < 10), key=lambda w: w[1])
                # Some reports leave Test blank. Anchor those rows on a numeric
                # Meas cell, while retaining result anchors for unreadable data.
                for word in candidates:
                    center = (word[0]+word[2])/2
                    if boundaries[1] < center < boundaries[2] and number(word[4]) is not None:
                        if not any(abs(word[1]-anchor[1]) <= 3 for anchor in results):
                            results.append(word)
                results.sort(key=lambda word: word[1])
                for result in results:
                    cells = [[] for _ in range(6)]
                    for word in sorted(candidates, key=lambda w: w[0]):
                        if abs(word[1]-result[1]) > 3:
                            continue
                        center = (word[0]+word[2])/2
                        col = sum(center > boundary for boundary in boundaries)
                        if center <= header[5][2]+15:
                            cells[col].append(word[4])
                    control, nom, measured, tolerance, deviation, test = [" ".join(c) for c in cells]
                    reading = number(measured)
                    if reading is None:
                        # Do not silently publish an incomplete recognized table.
                        raise ValueError(f"Unreadable feature-table measurement on page {page_index+1}, balloon {balloon[1]}: {measured!r}")
                    nominal = number(nom)
                    warnings = []
                    if not control or re.match(r"^[\d.+-]", control) and not re.match(r"\d+D\s+(Angle|Distance)", control, re.I):
                        kind = "Dimension" if nominal is not None else "Geometric control"
                        control = f"{kind} (symbol review): {control or 'not available in text'}"
                        warnings.append("Control symbol is graphical; verify its type against the PDF before using these statistics.")
                    if "\ufffd" in tolerance:
                        warnings.append("Tolerance sign is not encoded in the PDF text; verify the source and enter a signed tolerance.")
                    lower, upper = tolerance_limits(nominal, tolerance)
                    if lower is None or upper is None:
                        warnings.append("Specification limits require review; source tolerance retained without assuming a missing nominal or sign.")
                    if nominal is not None and number(deviation) is not None and abs(reading-nominal-number(deviation)) > number('0.002'):
                        warnings.append("Measurement, nominal and source deviation do not reconcile; verify the source.")
                    row = SummaryRow(feature[1].rstrip(', ') if feature else f"Balloon {balloon[1]}",
                                     control, nominal, lower, upper, tolerance, [reading],
                                     balloon[1], reference[1] if reference else "", str(page_index+1), test,
                                     f"Page {page_index+1}\n{context}\nControl | Nom | Meas | Tol | Dev | Test\n"+
                                     " | ".join(" ".join(c) for c in cells), warnings)
                    row.record_ids = [f"native_p{page_index+1}_b{balloon[1]}_r{len(rows)+1}"]
                    def preceding_label(label):
                        anchors = [w for w in words if w[4] == label and w[1] < marker[1]]
                        if not anchors:
                            return ''
                        anchor = max(anchors, key=lambda w: w[1])
                        return ' '.join(w[4] for w in sorted(words, key=lambda w: w[0])
                                        if abs(w[1]-anchor[1]) < 2 and w[0] > anchor[2])
                    alignment = preceding_label('Alignments') or active_alignment
                    coordinate = preceding_label('Systems') or active_coordinate
                    active_alignment, active_coordinate = alignment, coordinate
                    if alignment and coordinate:
                        row.group_context = ('page_context', alignment, coordinate)
                    rows.append(row)
    return rows
