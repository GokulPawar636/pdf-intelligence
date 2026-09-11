"""Combine independent PDF inspections without relying on row or page positions."""
from collections import Counter
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
from uuid import uuid4

import fitz
from openpyxl import Workbook

from app.reporting.summary_exporter import build_summary_rows, export_summary
from app.reporting.excel_exporter import _append, _style
from app.reporting.storage import save_workbook_atomic
from app.ingestion.pdf_loader import PDFLoader, save_raw_document
from app.structure.analyzer import analyze_file
from app.semantic.extractor import extract_file
from app.reporting.summary_exporter import measurement_values
from app.reporting.feature_tables import read_feature_tables
from app.ingestion.coverage import inspect_coverage


def source_details(path):
    with fitz.open(path) as pdf:
        text = "\n".join(page.get_text(sort=True) for page in pdf)
    def label(pattern):
        match = re.search(pattern + r":\s*([^\n]+)", text, re.I)
        return re.split(r"\s{2,}", match[1].strip())[0] if match else ""
    units = {u.lower() for u in re.findall(r"Length Units\s+(Millimeters|Millimetres|Inches|Centimeters|Meters)\b", text, re.I)}
    alignments = {" ".join(a.split()) for a in re.findall(r"Data Alignments\s+([^\n]+)", text, re.I)}
    return {"part": label(r"Part\s*(?:Number|No\.?|#)") or label("Part"),
            "name": label("Part name"), "date": label("Date"),
            "serial": label(r"Serial(?:\s*(?:No\.?|#))?"),
            "unit": {"millimeters": "mm", "millimetres": "mm", "inches": "in", "centimeters": "cm", "meters": "m"}.get(next(iter(units)), "") if len(units) == 1 else "",
            "alignment": next(iter(alignments)) if len(alignments) == 1 else "",
            "mixed_alignment": len(alignments) > 1}


def match_key(row, meta, file_index):
    def norm(value):
        return re.sub(r'\s*-\s*', '-', " ".join(str(value).casefold().split()))
    # Unknown part/feature/unit cannot justify joining records from different PDFs.
    per_feature_context = bool(row.group_context and row.group_context[0] == 'page_context')
    safe = bool(meta["part"] and row.control and row.object_name != "Not stated" and row.unit and
                (not meta.get("mixed_alignment") or per_feature_context))
    return (norm(meta["part"]), norm(row.object_name), norm(row.control), norm(row.balloon),
            norm(row.reference), norm(row.unit), row.nominal, row.lower, row.upper,
            norm(row.tolerance) if row.lower is None or row.upper is None else "",
            row.explicit_limits, row.group_context, '' if per_feature_context else norm(meta["alignment"]), None if safe else file_index)


def combine_rows(sources):
    headers, prepared, offset = [], [], 0
    for index, item in enumerate(sources):
        meta = item["details"]
        rows = deepcopy(item["rows"])
        for row in rows:
            if not row.unit and re.search(r"distance|diameter|radius|length|dimension|geometric control", row.control, re.I):
                row.unit = meta["unit"]
            per_feature_context = bool(row.group_context and row.group_context[0] == 'page_context')
            if len(sources) > 1 and (not meta["part"] or not row.unit or (meta.get("mixed_alignment") and not per_feature_context)):
                row.warnings.append("Missing part/unit or mixed alignment: kept separate across PDFs; verify the source before comparing.")
        width = max((len(row.measurements) for row in rows), default=1)
        for reading in range(width):
            title = f"PDF {index + 1}" + (f" / Reading {reading + 1}" if width > 1 else "")
            label = meta["date"] or "Date not stated"
            if meta["serial"]:
                label += f"\nSample {meta['serial']}"
            headers.append({"title": title, "label": label, "source": item["name"]})
        prepared.append((index, item, rows, offset))
        offset += width
    joined = {}
    for index, item, rows, offset in prepared:
        keys = [match_key(row, item["details"], index) for row in rows]
        counts = Counter(keys)
        for position, (row, key) in enumerate(zip(rows, keys)):
            if counts[key] > 1:
                key += (index, position)
                row.warnings.append("Duplicate feature identity in this PDF; kept separate to avoid an ambiguous match.")
            values = row.measurements
            if key not in joined:
                joined[key] = deepcopy(row)
                joined[key].measurements = [None] * len(headers)
                joined[key].page = ""
                joined[key].evidence = ""
                joined[key].record_ids = []
            target = joined[key]
            target.measurements[offset:offset + len(values)] = values
            target.warnings = list(dict.fromkeys(target.warnings + row.warnings))
            target.page += ("; " if target.page else "") + f"PDF {index + 1}: {row.page}"
            target.evidence += f"\nPDF {index + 1}: {item['name']}\n{row.evidence}\n"
            target.record_ids.extend(f"pdf{index + 1}:{rid}" for rid in row.record_ids)
            target.source_result = row.source_result
    multiple_parts = len({s["details"]["part"] for s in sources}) > 1
    for key, row in joined.items():
        if multiple_parts:
            row.object_name = f"{key[0] or 'Part not stated'} / {row.object_name}"
    return list(joined.values()), headers


def general_report(sources, output):
    """One readable source-field table instead of dozens of fragmented tabs."""
    book = Workbook()
    summary = book.active
    summary.title = "Summary Report"
    _append(summary, ["Source PDF", "Source Page", "Record", "Field", "Value", "Unit"])
    for source in sources:
        for record in source["payload"]["records"]:
            for field in record["fields"]:
                value = field["value"]
                raw = value.get("raw_value")
                raw = raw if raw is not None else value.get("normalized_value")
                parts = [raw[i:i+30000] for i in range(0, len(raw), 30000)] if isinstance(raw, str) and len(raw) > 30000 else [raw]
                for part in parts:
                    _append(summary, [source["name"], (value.get("source") or {}).get("page_number"),
                                      record.get("record_id"), field["name"].replace("_", " "), part, value.get("unit")])
    _style(summary, 1)
    differences = book.create_sheet("Differences")
    _append(differences, ["Report status", "Details"])
    _append(differences, ["Review required", "These PDFs do not provide a consistently identifiable measurement table. Source fields are retained in Summary Report; no statistical comparison has been invented."])
    _style(differences, 2)
    return save_workbook_atomic(book, output)


def run_batch(pdfs, output_dir="data/output", *, offline=True, strict=False,
              max_pages=500, max_mb=100, max_total_pages=2500, max_total_mb=500, progress=None):
    # Import locally to keep the CLI's optional --combine route free of cycles.
    from pipeline_main import _preflight, _write_json, select_report
    paths = [Path(path).expanduser().resolve() for path in pdfs]
    if not paths:
        raise ValueError("Choose at least one PDF.")
    total_pages, total_bytes = 0, 0
    for path in paths:
        _preflight(path, max_pages, max_mb)
        total_bytes += path.stat().st_size
        with fitz.open(path) as pdf:
            total_pages += len(pdf)
    if total_pages > max_total_pages or total_bytes > max_total_mb * 1024 * 1024:
        raise ValueError(f"This batch exceeds {max_total_pages} pages or {max_total_mb} MB. Split it into smaller batches.")
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "_" + uuid4().hex[:8]
    target = Path(output_dir).resolve() / "combined" / run_id
    target.mkdir(parents=True, exist_ok=False)
    manifest = {"status": "running", "run_id": run_id, "inputs": [], "warnings": [], "pages": total_pages}
    manifest_path = target / "manifest.json"
    _write_json(manifest_path, manifest)
    try:
        sources, seen = [], set()
        for index, path in enumerate(paths):
            if progress:
                progress(index, len(paths), path.name)
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            if digest in seen:
                manifest["warnings"].append(f"Duplicate PDF skipped: {path.name}")
                manifest["inputs"].append({"name": path.name, "sha256": digest, "duplicate": True})
                continue
            seen.add(digest)
            coverage = inspect_coverage(path)
            manifest["warnings"].extend(f"{path.name}: {note}" for note in coverage["warnings"])
            # Short per-source paths avoid Windows' path limit with descriptive filenames.
            work = target / "sources" / f"{index + 1:03d}"
            raw = PDFLoader(path, work, render_pages=False, extract_embedded_images=False).load()
            raw_path = save_raw_document(raw, work / "raw.json")
            layout = analyze_file(raw_path, work / "layout.json")
            # Prefer source-derived tables. AI is a fallback for unfamiliar
            # layouts, not a mandatory call for every page of a CMM report.
            semantic = extract_file(layout, work / "semantic.json", offline=True)
            payload = json.loads(semantic.read_text(encoding="utf-8"))
            structure = json.loads(layout.read_text(encoding="utf-8"))
            native_rows = read_feature_tables(path)
            strategy = 'local'
            if not offline and not native_rows and select_report(payload, structure) != 'summary':
                try:
                    interpreted = extract_file(layout, work / 'interpreted.json', offline=False,
                                               model=os.getenv('GROQ_MODEL', 'qwen/qwen3.8-27b'))
                    payload = json.loads(interpreted.read_text(encoding='utf-8'))
                    semantic = interpreted
                    strategy = 'groq'
                except Exception as exc:
                    # Retain local evidence, explicitly mark degraded output.
                    payload.setdefault('warnings', []).append(
                        f'Groq interpretation unavailable ({type(exc).__name__}); report uses local extraction and requires review.')
                    strategy = 'local_after_ai_failure'
            source_warnings = list(payload.get("warnings", []))
            for record in payload["records"]:
                source_warnings.extend(record.get("warnings", []))
                for item in record["fields"]:
                    value = item["value"]
                    candidate = value.get("raw_value") if value.get("raw_value") is not None else value.get("normalized_value")
                    try:
                        measurement_values({re.sub(r"[^\w]+", "_", item["name"].casefold()).strip("_"): [candidate]})
                    except ValueError as exc:
                        source_warnings.append(str(exc))
            summary_rows = build_summary_rows(payload, structure)
            is_measurement = select_report(payload, structure) == "summary"
            if not is_measurement:
                if native_rows:
                    summary_rows = native_rows
                    is_measurement = True
            sources.append({"name": path.name, "details": source_details(path), "payload": payload,
                            "rows": summary_rows, "semantic": semantic, "layout": layout,
                            "is_measurement": is_measurement})
            manifest["inputs"].append({"name": path.name, "sha256": digest, "pages": raw.document.page_count, "duplicate": False, "coverage": coverage, 'extraction_strategy': strategy})
            manifest["warnings"].extend(f"{path.name}: {warning}" for warning in source_warnings)
            _write_json(manifest_path, manifest)
        if not any(source["payload"]["records"] for source in sources):
            raise ValueError("No readable data was found. Scanned PDFs need OCR before conversion.")
        rows, headers = combine_rows(sources)
        use_measurements = bool(rows) and all(source["is_measurement"] for source in sources)
        if not use_measurements:
            manifest["warnings"].append("A comparable measurement layout could not be identified in every PDF. A consolidated source-field report was produced for review.")
        manifest["warnings"].extend(note for row in rows for note in row.warnings)
        manifest["warnings"] = list(dict.fromkeys(manifest["warnings"]))
        if strict and manifest["warnings"]:
            raise ValueError("Quality checks blocked export: " + " | ".join(manifest["warnings"][:5]) +
                             (f" | {len(manifest['warnings'])-5} additional review notes." if len(manifest["warnings"]) > 5 else ""))
        output = target / "Combined_Report.xlsx"
        if use_measurements:
            parts = sorted({s["details"]["part"] for s in sources if s["details"]["part"]})
            title = "Inspection Summary" + (f" - {parts[0]}" if len(parts) == 1 else "")
            metadata = [("Part", ", ".join(parts) or "Not stated"), ("PDFs", str(len(sources))),
                        ("Readings", str(sum(v is not None for row in rows for v in row.measurements)))]
            export_summary(sources[0]["semantic"], sources[0]["layout"], output, summary_rows=rows,
                           measurement_headers=headers, report_title=title,
                           source_caption=('REVIEW REQUIRED. Image-only content or source interpretation may be incomplete. ' if manifest['warnings'] else '') + f"{len(sources)} PDFs / {sum(i['pages'] for i in manifest['inputs'] if not i['duplicate'])} pages. Columns follow upload order. Blank cells mean no matching reading in that PDF.",
                           title_fields_override=metadata)
        else:
            general_report(sources, output)
        manifest.update(status="needs_review" if manifest["warnings"] else "complete",
                        accuracy_verified=False, checks_passed=not bool(manifest["warnings"]),
                        selected_report="summary" if use_measurements else "general",
                        measurement_columns=len(headers) if use_measurements else 0,
                        measurement_rows=len(rows) if use_measurements else 0,
                        output=str(output), output_sha256=hashlib.sha256(output.read_bytes()).hexdigest())
        _write_json(manifest_path, manifest)
        if progress:
            progress(len(paths), len(paths), "Report ready")
        return output
    except Exception as exc:
        manifest.update(status="failed", error_type=type(exc).__name__, error=str(exc))
        _write_json(manifest_path, manifest)
        raise
