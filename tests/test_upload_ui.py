import json
from types import SimpleNamespace

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from streamlit_app import preview_html


def test_uploads_rebuild_automatically_and_sheet_changes_reuse_result(tmp_path, monkeypatch):
    book = Workbook()
    book.active.title = "Summary Report"
    book.active["A1"] = "<script>not executable</script>"
    book.active["B1"] = 18.081
    book.active["B1"].number_format = "0.000"
    book.create_sheet("Differences")["A1"] = "Difference"
    output = tmp_path / "report.xlsx"
    book.save(output)
    output.with_name("manifest.json").write_text(json.dumps({"warnings": [], "pages": 2, "measurement_columns": 2}))
    selected = [SimpleNamespace(name=f"input{i}.pdf", getvalue=lambda: b"test PDF") for i in range(2)]
    calls = []
    def batch(paths, *args, **kwargs):
        calls.append(len(paths))
        return output
    monkeypatch.setattr("streamlit.file_uploader", lambda *args, **kwargs: selected)
    monkeypatch.setattr("app.batch.run_batch", batch)
    app = AppTest.from_file("streamlit_app.py", default_timeout=30).run()
    assert not app.exception
    assert len(app.checkbox) == 0
    assert app.radio[0].options == ['PDF', 'Excel']
    assert calls == [2]
    assert len(app.get("download_button")) == 1
    app.selectbox[0].select("Differences").run()
    assert not app.exception and calls == [2]
    selected.append(SimpleNamespace(name="third.pdf", getvalue=lambda: b"third PDF"))
    app.run()
    assert not app.exception and calls == [2, 3]
    html = preview_html(output.read_bytes(), "Summary Report")
    assert "&lt;script&gt;" in html and "<script>" not in html
    assert "18.081" in html
