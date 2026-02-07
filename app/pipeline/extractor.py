"""Step 1: PDF to text extraction (native + OCR for scanned documents)."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf
from pdf2image import convert_from_path
from pytesseract import image_to_string

logger = logging.getLogger(__name__)

MIN_CHARS_FOR_NATIVE = 50  # per page – below this we assume scanned


@dataclass
class PageText:
    page_num: int
    text: str


@dataclass
class ExtractionResult:
    pages: list[PageText] = field(default_factory=list)
    is_scanned: bool = False
    total_text: str = ""

    def build_total_text(self) -> None:
        self.total_text = "\n\n".join(p.text for p in self.pages if p.text.strip())


def _extract_native(pdf_path: Path) -> ExtractionResult | None:
    """Try native text extraction with PyMuPDF."""
    doc = pymupdf.open(str(pdf_path))
    pages: list[PageText] = []
    has_text = False

    for i, page in enumerate(doc):
        text = page.get_text().strip()
        pages.append(PageText(page_num=i + 1, text=text))
        if len(text) >= MIN_CHARS_FOR_NATIVE:
            has_text = True

    doc.close()

    if not has_text:
        return None

    result = ExtractionResult(pages=pages, is_scanned=False)
    result.build_total_text()
    return result


def _extract_ocr(pdf_path: Path) -> ExtractionResult:
    """Fall back to OCR for scanned PDFs."""
    logger.info("Native extraction insufficient – running OCR on %s", pdf_path.name)
    images = convert_from_path(str(pdf_path), dpi=300)
    pages: list[PageText] = []

    for i, image in enumerate(images):
        text = image_to_string(image).strip()
        pages.append(PageText(page_num=i + 1, text=text))

    result = ExtractionResult(pages=pages, is_scanned=True)
    result.build_total_text()
    return result


def extract_text(pdf_path: str | Path) -> ExtractionResult:
    """Extract text from a PDF file. Tries native extraction first, falls back to OCR."""
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    result = _extract_native(pdf_path)
    if result is not None:
        logger.info("Native extraction succeeded for %s (%d pages)", pdf_path.name, len(result.pages))
        return result

    return _extract_ocr(pdf_path)
