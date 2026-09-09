"""Simple local PDF upload, workbook preview, and download interface."""
from __future__ import annotations

from io import BytesIO
from html import escape
import json
import logging
from pathlib import Path
import tempfile

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from pipeline_main import run_pipeline


def clear_result() -> None:
    st.session_state.pop("result", None)


def convert_upload(name: str, contents: bytes) -> dict:
    if not contents:
        raise ValueError("The uploaded file is empty. Please choose another PDF.")
    if len(contents) > 100 * 1024 * 1024:
        raise ValueError("Please choose a PDF smaller than 100 MB.")
    safe_name = Path(name.replace("\\", "/")).name
    if Path(safe_name).suffix.lower() != ".pdf":
        raise ValueError("Please choose a PDF file.")
    with tempfile.TemporaryDirectory(prefix="pdf-intelligence-") as workspace:
        source = Path(workspace) / safe_name
        source.write_bytes(contents)
        output = run_pipeline(source, Path(workspace) / "output", offline=True, report="auto")
        manifest = json.loads(output.with_name("manifest.json").read_text(encoding="utf-8"))
        return {"name": output.name, "contents": output.read_bytes(),
                "warnings": manifest.get("warnings", []), "source": safe_name}


def main() -> None:
    st.set_page_config(page_title="PDF to Excel", page_icon=":page_facing_up:", layout="wide")
    st.title("PDF to Excel")
    st.caption("Upload a PDF, preview the extracted data, and download your Excel file.")
    uploaded = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_upload", on_change=clear_result)
    st.caption("Up to 100 MB and 500 pages. Files are processed locally.")

    if st.button("Convert to Excel", type="primary", disabled=uploaded is None):
        clear_result()
        with st.spinner("Reading your PDF and preparing Excel..."):
            try:
                st.session_state["result"] = convert_upload(uploaded.name, uploaded.getvalue())
            except ValueError as exc:
                st.error(str(exc))
            except Exception:
                logging.getLogger(__name__).exception("PDF conversion failed")
                st.error("We couldn't convert this PDF. Please check that it opens correctly and contains readable text.")

    result = st.session_state.get("result")
    if result is None:
        return
    st.divider()
    if result["warnings"]:
        st.warning("Your Excel file is ready, but some extracted data needs review.")
        with st.expander("View review notes"):
            for warning in result["warnings"]:
                st.write(warning)
    else:
        st.success("Your Excel file is ready.")
    st.download_button("Download Excel", result["contents"], file_name=result["name"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    st.subheader("Output preview")
    st.caption(result["source"])
    workbook = load_workbook(BytesIO(result["contents"]), read_only=True, data_only=True)
    try:
        selected = st.selectbox("Worksheet", workbook.sheetnames)
        sheet = workbook[selected]
        rows = [["" if value is None else str(value) for value in row]
                for row in sheet.iter_rows(max_row=min(sheet.max_row, 200), values_only=True)]
        # Preserve title rows and multi-row headers exactly as worksheet cells.
        headers = "<th>Row</th>" + "".join(
            f"<th>{get_column_letter(i)}</th>" for i in range(1, sheet.max_column + 1))
        body = "".join(
            f"<tr><th>{number}</th>" + "".join(f"<td>{escape(value)}</td>" for value in row) + "</tr>"
            for number, row in enumerate(rows, 1))
        st.html(
            "<style>.excel-preview{overflow:auto;max-height:450px}"
            ".excel-preview table{border-collapse:collapse;font-size:14px}"
            ".excel-preview th,.excel-preview td{border:1px solid #ccc;padding:8px;"
            "min-width:90px;max-width:450px;overflow-wrap:anywhere}"
            ".excel-preview thead th{position:sticky;top:0;background:#e8eef5;color:#172638}"
            "</style><div class='excel-preview'><table><thead><tr>" + headers +
            "</tr></thead><tbody>" + body + "</tbody></table></div>")
        st.caption("Preview shows the first 200 rows per worksheet. Download Excel for the full formatted report.")
    finally:
        workbook.close()


if __name__ == "__main__":
    main()
