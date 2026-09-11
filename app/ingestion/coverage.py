"""Page-level extraction checks; these are not an accuracy certification."""
import fitz


def inspect_coverage(path):
    pages, warnings = [], []
    with fitz.open(path) as document:
        for index, page in enumerate(document, 1):
            text = page.get_text().strip()
            area = max(page.rect.get_area(), 1)
            largest_image = max((fitz.Rect(image['bbox']).get_area() / area
                                 for image in page.get_image_info()), default=0)
            status = "native_text"
            if len(text) < 20:
                if largest_image >= .25:
                    status = "ocr_required"
                    warnings.append(f"Page {index}: image content has no usable text. OCR is required; its contents are not included in extracted measurements.")
                elif page.get_drawings():
                    status = "drawing_review"
                    warnings.append(f"Page {index}: drawing-only content requires visual review and cannot be certified from text extraction.")
                else:
                    status = "blank_or_sparse"
            elif largest_image >= .7:
                status = "image_text_review"
                warnings.append(f"Page {index}: a large image accompanies text. Verify OCR/text coverage against the page image.")
            if '\ufffd' in text:
                warnings.append(f"Page {index}: unreadable text characters detected. Verify symbols and values against the PDF.")
            pages.append({"page": index, "status": status, "text_characters": len(text)})
    return {"pages": pages, "warnings": warnings, "accuracy_verified": False}
