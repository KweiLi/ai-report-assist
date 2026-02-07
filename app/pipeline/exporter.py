"""Step 6: Export debiased report to PDF or DOCX."""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt, RGBColor
from fpdf import FPDF

logger = logging.getLogger(__name__)


def export_docx(
    text: str,
    output_path: str | Path,
    title: str = "Debiased Report",
    entities_found: list[dict] | None = None,
    changes_summary: str | None = None,
    acronyms_preserved: list[dict] | None = None,
) -> Path:
    """Export debiased report with masking, debiasing, and acronym details to Word."""
    output_path = Path(output_path)
    doc = Document()

    # Title
    heading = doc.add_heading(title, level=1)
    heading.runs[0].font.size = Pt(16)

    # Metadata
    meta = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    meta.runs[0].font.color.rgb = RGBColor(128, 128, 128)
    meta.runs[0].font.size = Pt(9)

    # --- Section 1: Debiased Report ---
    doc.add_heading("Debiased Report", level=2)
    for paragraph in text.split("\n"):
        if paragraph.strip():
            doc.add_paragraph(paragraph)

    # --- Section 2: Debiasing Details ---
    if changes_summary:
        doc.add_page_break()
        doc.add_heading("Debiasing Details", level=2)
        intro = doc.add_paragraph("Summary of changes made to reduce bias in the report:")
        intro.runs[0].font.italic = True
        intro.runs[0].font.size = Pt(9)
        for line in changes_summary.split("\n"):
            if line.strip():
                doc.add_paragraph(line)

    # --- Section 3: Masking Details ---
    if entities_found:
        doc.add_page_break()
        doc.add_heading("Masking Details", level=2)
        intro = doc.add_paragraph(
            f"{len(entities_found)} sensitive entities were detected and masked "
            "before sending the report to the AI for debiasing. "
            "Only masked text was transmitted to the cloud."
        )
        intro.runs[0].font.italic = True
        intro.runs[0].font.size = Pt(9)

        table = doc.add_table(rows=1, cols=4)
        table.style = "Light Grid Accent 1"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = table.rows[0].cells
        for i, label in enumerate(["Type", "Original Value", "Masked Token", "Confidence"]):
            hdr[i].text = label
            hdr[i].paragraphs[0].runs[0].font.bold = True
            hdr[i].paragraphs[0].runs[0].font.size = Pt(9)

        for entity in entities_found:
            row = table.add_row().cells
            row[0].text = entity["entity_type"]
            row[1].text = entity["original"]
            row[2].text = entity["token"]
            row[3].text = str(entity["score"])
            for cell in row:
                cell.paragraphs[0].runs[0].font.size = Pt(9)

    # --- Section 4: Acronyms & Abbreviations ---
    if acronyms_preserved:
        doc.add_page_break()
        doc.add_heading("Acronyms & Abbreviations", level=2)
        intro = doc.add_paragraph(
            f"{len(acronyms_preserved)} acronyms/abbreviations were identified in the report. "
            "These are standard law enforcement terminology and were intentionally "
            "preserved (not masked) during processing."
        )
        intro.runs[0].font.italic = True
        intro.runs[0].font.size = Pt(9)

        table = doc.add_table(rows=1, cols=3)
        table.style = "Light Grid Accent 1"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = table.rows[0].cells
        for i, label in enumerate(["Acronym", "Initially Detected As", "Action"]):
            hdr[i].text = label
            hdr[i].paragraphs[0].runs[0].font.bold = True
            hdr[i].paragraphs[0].runs[0].font.size = Pt(9)

        for acr in acronyms_preserved:
            row = table.add_row().cells
            row[0].text = acr["text"]
            row[1].text = acr["detected_as"]
            row[2].text = acr["reason"]
            for cell in row:
                cell.paragraphs[0].runs[0].font.size = Pt(9)

    doc.save(str(output_path))
    logger.info("Exported DOCX to %s", output_path)
    return output_path


def export_pdf(
    text: str,
    output_path: str | Path,
    title: str = "Debiased Report",
    entities_found: list[dict] | None = None,
    changes_summary: str | None = None,
    acronyms_preserved: list[dict] | None = None,
) -> Path:
    """Export debiased report with masking, debiasing, and acronym details to PDF."""
    output_path = Path(output_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)

    # --- Page 1: Debiased Report ---
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(128, 128, 128)
    pdf.cell(0, 6, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    _pdf_section_heading(pdf, "Debiased Report")
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 10)
    for paragraph in text.split("\n"):
        if paragraph.strip():
            pdf.multi_cell(0, 5, paragraph)
            pdf.ln(2)

    # --- Page 2: Debiasing Details ---
    if changes_summary:
        pdf.add_page()
        _pdf_section_heading(pdf, "Debiasing Details")
        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(0, 5, "Summary of changes made to reduce bias in the report:")
        pdf.ln(3)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "", 10)
        for line in changes_summary.split("\n"):
            if line.strip():
                pdf.multi_cell(0, 5, line)
                pdf.ln(2)

    # --- Page 3: Masking Details ---
    if entities_found:
        pdf.add_page()
        _pdf_section_heading(pdf, "Masking Details")
        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(
            0, 5,
            f"{len(entities_found)} sensitive entities were detected and masked "
            "before sending the report to the AI for debiasing. "
            "Only masked text was transmitted to the cloud.",
        )
        pdf.ln(4)
        pdf.set_text_color(0, 0, 0)

        col_widths = [45, 55, 55, 25]
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(230, 230, 230)
        for w, label in zip(col_widths, ["Type", "Original Value", "Masked Token", "Confidence"]):
            pdf.cell(w, 7, label, border=1, fill=True)
        pdf.ln()

        pdf.set_font("Helvetica", "", 8)
        for entity in entities_found:
            row_data = [
                entity["entity_type"],
                _truncate(entity["original"], 30),
                entity["token"],
                str(entity["score"]),
            ]
            for w, val in zip(col_widths, row_data):
                pdf.cell(w, 6, val, border=1)
            pdf.ln()

    # --- Page 4: Acronyms & Abbreviations ---
    if acronyms_preserved:
        pdf.add_page()
        _pdf_section_heading(pdf, "Acronyms & Abbreviations")
        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(
            0, 5,
            f"{len(acronyms_preserved)} acronyms/abbreviations were identified in the report. "
            "These are standard law enforcement terminology and were intentionally "
            "preserved (not masked) during processing.",
        )
        pdf.ln(4)
        pdf.set_text_color(0, 0, 0)

        col_widths = [30, 50, 100]
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(230, 230, 230)
        for w, label in zip(col_widths, ["Acronym", "Detected As", "Action"]):
            pdf.cell(w, 7, label, border=1, fill=True)
        pdf.ln()

        pdf.set_font("Helvetica", "", 8)
        for acr in acronyms_preserved:
            pdf.cell(30, 6, acr["text"], border=1)
            pdf.cell(50, 6, acr["detected_as"], border=1)
            pdf.cell(100, 6, _truncate(acr["reason"], 55), border=1)
            pdf.ln()

    pdf.output(str(output_path))
    logger.info("Exported PDF to %s", output_path)
    return output_path


def _pdf_section_heading(pdf: FPDF, text: str) -> None:
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(30, 30, 60)
    pdf.cell(0, 9, text, new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(30, 30, 60)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(4)


def _truncate(text: str, max_len: int) -> str:
    return text if len(text) <= max_len else text[: max_len - 3] + "..."
