from io import BytesIO
import fitz
from PIL import Image, ImageDraw
import pytest
from app.ingestion.coverage import inspect_coverage
from app.batch import run_batch


def mixed_pdf(path):
    document = fitz.open()
    document.new_page().insert_text((40, 40), "Inspection notes: native text is readable.")
    image = Image.new('RGB', (600, 800), 'white')
    ImageDraw.Draw(image).text((40, 40), 'Measurement 12.345', fill='black')
    buffer = BytesIO()
    image.save(buffer, format='PNG')
    page = document.new_page()
    page.insert_image(page.rect, stream=buffer.getvalue())
    document.new_page()
    document.save(path)
    document.close()
    return path


def test_mixed_pdf_does_not_hide_scanned_page(tmp_path):
    result = inspect_coverage(mixed_pdf(tmp_path / 'mixed.pdf'))
    assert [p['status'] for p in result['pages']] == ['native_text', 'ocr_required', 'blank_or_sparse']
    assert 'Page 2' in result['warnings'][0]
    assert result['accuracy_verified'] is False


def test_strict_batch_blocks_incomplete_coverage(tmp_path):
    with pytest.raises(ValueError, match='OCR is required'):
        run_batch([mixed_pdf(tmp_path / 'mixed.pdf')], tmp_path / 'out', strict=True)
    assert not list((tmp_path / 'out').rglob('*.xlsx'))
