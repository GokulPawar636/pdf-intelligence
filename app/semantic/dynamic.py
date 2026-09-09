"""Evidence-preserving extraction without a document-type template."""
from __future__ import annotations

import re
from typing import Any

from app.models.semantic import SemanticField, SemanticRecord, SemanticValue, SourceReference


def field_name(text: str, fallback: str) -> str:
    return re.sub(r"[^\w]+", "_", text.casefold(), flags=re.UNICODE).strip("_") or fallback


def normalize_value(text: str, name: str = "") -> Any:
    """Only normalize unambiguous numbers; identifiers and locale formats stay exact."""
    if re.search(r"(^|_)(id|code|zip|postal|phone|account|sku|number)($|_)", name):
        return text
    if not re.fullmatch(r"[+-]?(?:0|[1-9]\d*)(?:\.\d+)?", text):
        return text
    if len(re.sub(r"\D", "", text)) > 15:
        return text
    return float(text) if "." in text else int(text)


def extract_tables(document: dict[str, Any], source_factory) -> list[SemanticRecord]:
    records = []
    for page in document.get("pages", []):
        page_number = int(page.get("page_number", 0))
        for table_index, table in enumerate(page.get("tables", []), 1):
            table_id = str(table.get("table_id") or f"p{page_number:03d}_t{table_index:03d}")
            rows: dict[int, list[dict]] = {}
            for cell in table.get("cells", []):
                rows.setdefault(int(cell.get("row", 0)), []).append(cell)
            ordered = sorted(rows)
            if not ordered:
                continue
            columns = sorted({int(c.get("column", 0)) for cells in rows.values() for c in cells})
            first = {int(c.get("column", 0)): str(c.get("text", "")).strip() for c in rows[ordered[0]]}
            # Header inference remains an explicit heuristic. Numeric first rows
            # are always retained, and uncertain all-text tables lose no rows.
            explicit = table.get("evidence", {}).get("has_header")
            has_header = explicit if isinstance(explicit, bool) else (
                len(ordered) > 1
                and len(first) == len(columns)
                and all(v and any(ch.isalpha() for ch in v) for v in first.values())
                and not any(v.endswith(":") for v in first.values())
                and any(
                    isinstance(normalize_value(str(c.get("text", "")).strip()), (int, float))
                    for row in ordered[1:] for c in rows[row]
                )
            )
            headers: dict[int, str] = {}
            used: set[str] = set()
            for column in columns:
                base = field_name(first.get(column, ""), f"column_{column + 1}") if has_header else f"column_{column + 1}"
                name, suffix = base, 2
                while name in used:
                    name = f"{base}_{suffix}"
                    suffix += 1
                headers[column] = name
                used.add(name)
            confidence = max(0.0, min(1.0, float(table.get("confidence", 0.7))))
            for row in ordered[1:] if has_header else ordered:
                fields = []
                for cell in sorted(rows[row], key=lambda c: int(c.get("column", 0))):
                    raw = str(cell.get("text", "")).strip()
                    name = headers[int(cell.get("column", 0))]
                    source = source_factory(page_number, table_id, cell)
                    source.confidence = confidence
                    fields.append(SemanticField(name=name, value=SemanticValue(
                        raw_value=raw, normalized_value=normalize_value(raw, name), source=source,
                    )))
                if any(f.value.raw_value for f in fields):
                    records.append(SemanticRecord(
                        record_id=f"{table_id}_r{row}", record_type="table_row", fields=fields,
                        source_pages=[page_number], confidence=confidence,
                        warnings=["First row interpreted as column headers; verify against source."] if has_header else [
                            "Column headers are uncertain; all rows retained with positional column names."
                        ],
                    ))
    return records


def extract_text(document: dict[str, Any], region_filter) -> list[SemanticRecord]:
    """Keep labeled values and prose when semantic interpretation is unavailable."""
    records = []
    for page in document.get("pages", []):
        page_number = int(page.get("page_number", 0))
        for index, region in enumerate(region_filter(page), 1):
            text = str(region.get("text", "")).strip()
            if not text:
                continue
            region_id = str(region.get("region_id") or f"p{page_number:03d}_text{index:03d}")
            fields, remaining = [], []
            for line in text.splitlines():
                match = re.fullmatch(r"\s*([^:\n]{1,80}):\s+(.+)", line)
                if match and any(c.isalpha() for c in match[1]):
                    label, raw = match.groups()
                    name = field_name(label, "value")
                    fields.append(SemanticField(name=name, value=SemanticValue(
                        raw_value=raw, normalized_value=normalize_value(raw, name),
                        source=SourceReference(page_number=page_number, source_text=line,
                            bbox=region.get("bbox"), source_id=region_id,
                            extraction_method="labeled_text", confidence=0.8),
                    )))
                else:
                    remaining.append(line)
            if remaining:
                raw = "\n".join(remaining)
                fields.append(SemanticField(name="text", value=SemanticValue(
                    raw_value=raw, normalized_value=raw,
                    source=SourceReference(page_number=page_number, source_text=raw,
                        bbox=region.get("bbox"), source_id=region_id,
                        extraction_method="source_text", confidence=0.8),
                )))
            records.append(SemanticRecord(record_id=region_id,
                record_type="labeled_values" if any(f.name != "text" for f in fields) else "document_text",
                fields=fields, source_pages=[page_number], confidence=0.8))
    return records
