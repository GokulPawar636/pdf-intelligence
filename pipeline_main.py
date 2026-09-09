"""Validated, isolated PDF-to-Excel runs with an auditable completion manifest."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import logging
import os
from pathlib import Path
import re
from uuid import uuid4

import fitz

from app.ingestion.pdf_loader import PDFLoader, save_raw_document
from app.reporting.excel_exporter import export_semantic_to_excel
from app.reporting.summary_exporter import export_summary, build_summary_rows, measurement_values
from app.semantic.extractor import extract_file
from app.structure.analyzer import analyze_file

logger = logging.getLogger(__name__)


def _write_json(path: Path, payload: dict) -> None:
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def select_report(payload: dict, layout: dict) -> str:
    rows = build_summary_rows(payload, layout)
    if not rows or any(not row.control or row.object_name == "Not stated" for row in rows):
        return "dynamic"
    for record in payload.get("records", []):
        if record["record_type"] == "table_row":
            positional = all(re.fullmatch(r"column_\d+", f["name"]) for f in record["fields"])
            if not positional and not build_summary_rows({"records": [record]}, layout):
                return "dynamic"
    return "summary"


def _preflight(source: Path, max_pages: int, max_mb: int) -> None:
    if max_pages < 1 or max_mb < 1:
        raise ValueError("PDF size and page limits must be positive.")
    if not source.is_file():
        raise FileNotFoundError(f"Input PDF not found: {source}")
    if source.suffix.casefold() != ".pdf":
        raise ValueError("Input must be a .pdf file.")
    if source.stat().st_size > max_mb * 1024 * 1024:
        raise ValueError(f"PDF exceeds the configured {max_mb} MB limit.")
    with fitz.open(source) as pdf:
        if not pdf.is_pdf:
            raise ValueError("File content is not a PDF.")
        if pdf.needs_pass:
            raise ValueError("Password-protected PDFs must be decrypted before processing.")
        if not 1 <= len(pdf) <= max_pages:
            raise ValueError(f"PDF must contain between 1 and {max_pages} pages.")


def run_pipeline(pdf: str | Path, output_dir: str | Path = "data/output", *,
                 offline: bool = False, model: str = "qwen/qwen3.8-27b", report: str = "auto",
                 strict: bool = False, max_pages: int = 500, max_mb: int = 100,
                 render_assets: bool = False) -> Path:
    if report not in {"auto", "dynamic", "summary"}:
        raise ValueError("report must be auto, dynamic or summary")
    source = Path(pdf).expanduser().resolve()
    _preflight(source, max_pages, max_mb)
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "_" + uuid4().hex[:8]
    target = Path(output_dir).expanduser().resolve() / source.stem / run_id
    target.mkdir(parents=True, exist_ok=False)
    intermediate = target / "intermediate"
    manifest_path = target / "manifest.json"
    manifest = {"run_id": run_id, "status": "running", "source": str(source),
                "requested_report": report, "offline": offline, "stage": "ingestion"}
    _write_json(manifest_path, manifest)
    try:
        raw = PDFLoader(source, intermediate, render_pages=render_assets,
                        extract_embedded_images=render_assets).load()
        raw_path = save_raw_document(raw, intermediate / "raw_document.json")
        manifest.update(source_sha256=raw.document.sha256, pages=raw.document.page_count, stage="structure")
        _write_json(manifest_path, manifest)
        structure = analyze_file(raw_path, intermediate / "structured_document.json")
        manifest["stage"] = "semantic"
        _write_json(manifest_path, manifest)
        semantic = extract_file(structure, intermediate / "semantic_document.json", model=model, offline=offline)
        payload = json.loads(Path(semantic).read_text(encoding="utf-8"))
        layout = json.loads(Path(structure).read_text(encoding="utf-8"))
        issues = list(payload.get("warnings", []))
        # Record-level uncertainty must also reach the manifest and upload UI.
        # Keeping the raw cells does not establish that their inferred headers are correct.
        issues.extend(f"Record {record.get('record_id')}: {warning}"
                      for record in payload["records"] for warning in record.get("warnings", []))
        if not payload["records"]:
            issues.append("No readable records were extracted; OCR or manual review may be required.")
        for record in payload["records"]:
            for item in record["fields"]:
                name = re.sub(r"[^\w]+", "_", item["name"].casefold()).strip("_")
                value = item["value"]
                raw_value = value.get("raw_value")
                if raw_value is None:
                    raw_value = value.get("normalized_value")
                try:
                    measurement_values({name: [raw_value]})
                except ValueError as exc:
                    issues.append(f"Record {record.get('record_id')}: {exc}")
        summary_rows = build_summary_rows(payload, layout)
        issues.extend(note for row in summary_rows for note in row.warnings)
        issues = list(dict.fromkeys(issues))
        manifest.update(records=len(payload["records"]), warnings=issues,
                        quality={"warning_count": len(issues),
                                 "measurement_rows": len(summary_rows),
                                 "measurement_readings": sum(len(row.measurements) for row in summary_rows),
                                 "review_required": bool(issues),
                                 "accuracy_verified": False})
        if strict and issues:
            raise ValueError("Strict quality check failed. " + "; ".join(issues))
        payload["warnings"] = issues
        _write_json(Path(semantic), payload)
        chosen = select_report(payload, layout) if report == "auto" else report
        manifest.update(selected_report=chosen, stage="export")
        _write_json(manifest_path, manifest)
        if chosen == "summary":
            output = export_summary(semantic, structure, target / f"{source.stem}_summary.xlsx")
        else:
            output = export_semantic_to_excel(semantic, target / f"{source.stem}.xlsx")
        manifest.update(status="needs_review" if issues else "complete", stage="complete",
                        output=str(output), output_sha256=hashlib.sha256(output.read_bytes()).hexdigest())
        _write_json(manifest_path, manifest)
        return output
    except Exception as exc:
        manifest.update(status="failed", error_type=type(exc).__name__)
        _write_json(manifest_path, manifest)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert PDFs to validated, document-driven Excel workbooks.")
    parser.add_argument("pdf", nargs="+", help="One or more input PDF paths; each receives an isolated workbook")
    parser.add_argument("--output-dir", default="data/output")
    parser.add_argument("--offline", action="store_true", help="Extract locally without sending content to Groq")
    parser.add_argument("--model", default="qwen/qwen3.8-27b")
    parser.add_argument("--report", choices=["auto", "dynamic", "summary"], default="auto")
    parser.add_argument("--strict", action="store_true", help="Reject empty extraction or document/measurement quality warnings")
    parser.add_argument("--max-pages", type=int, default=500)
    parser.add_argument("--max-mb", type=int, default=100)
    parser.add_argument("--render-assets", action="store_true", help="Save page and embedded images for diagnosis")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    failed = False
    for source in args.pdf:
        try:
            output = run_pipeline(source, args.output_dir, offline=args.offline, model=args.model,
                                  report=args.report, strict=args.strict, max_pages=args.max_pages,
                                  max_mb=args.max_mb, render_assets=args.render_assets)
            print(f"Workbook: {output}")
        except Exception as exc:
            logger.error("%s: %s", Path(source).name, exc)
            failed = True
    return int(failed)


if __name__ == "__main__":
    raise SystemExit(main())
