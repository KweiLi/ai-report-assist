"""Step 6: Export debiased report to PDF or DOCX."""

from __future__ import annotations

import logging
from collections import Counter
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from fpdf import FPDF

from app.pipeline.unmasker import unmask_phrase

logger = logging.getLogger(__name__)

# Bias type display colours (hex without #) – mirrors BIAS_COLORS in debiaser.py
_BIAS_HEX = {
    "RACIAL_ETHNIC": "e74c3c",
    "GENDER": "e67e22",
    "SOCIOECONOMIC": "9b59b6",
    "STEREOTYPING": "f39c12",
    "INFLAMMATORY": "c0392b",
    "CONFIRMATION": "3498db",
    "SUBJECTIVE": "e84393",
}

_BIAS_LABELS = {
    "RACIAL_ETHNIC": "Racial / Ethnic",
    "GENDER": "Gender",
    "SOCIOECONOMIC": "Socioeconomic",
    "STEREOTYPING": "Stereotyping",
    "INFLAMMATORY": "Inflammatory",
    "CONFIRMATION": "Confirmation",
    "SUBJECTIVE": "Subjective",
}


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _lighten_hex(hex_color: str, factor: float = 0.85) -> str:
    """Lighten a hex colour for use as a background tint."""
    r, g, b = _hex_to_rgb(hex_color)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"{r:02x}{g:02x}{b:02x}"


# ── DOCX helpers ──

def _set_cell_shading(cell, hex_color: str) -> None:
    """Set background shading colour on a DOCX table cell."""
    tc_pr = cell._element.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:shd"))
    if existing is not None:
        tc_pr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color.upper())
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def _set_cell_border_left(cell, hex_color: str, width: int = 12) -> None:
    """Set a coloured left border on a DOCX table cell."""
    tc_pr = cell._element.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(width))
    left.set(qn("w:color"), hex_color.upper())
    left.set(qn("w:space"), "0")
    existing = borders.find(qn("w:left"))
    if existing is not None:
        borders.remove(existing)
    borders.append(left)


def _hide_cell_borders(cell) -> None:
    """Remove all borders except left from a DOCX table cell."""
    tc_pr = cell._element.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for side in ("w:top", "w:right", "w:bottom"):
        el = OxmlElement(side)
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:color"), "FFFFFF")
        existing = borders.find(qn(side))
        if existing is not None:
            borders.remove(existing)
        borders.append(el)


def _docx_bias_summary_table(doc: Document, bias_changes: list[dict]) -> None:
    """Add a colour-coded summary table of bias type counts."""
    counts = Counter(c["bias_type"] for c in bias_changes)
    table = doc.add_table(rows=1, cols=len(counts))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (bias_type, count) in enumerate(counts.most_common()):
        cell = table.cell(0, i)
        color = _BIAS_HEX.get(bias_type, "999999")
        _set_cell_shading(cell, _lighten_hex(color, 0.75))
        _set_cell_border_left(cell, color, width=18)
        p = cell.paragraphs[0]
        label_run = p.add_run(_BIAS_LABELS.get(bias_type, bias_type))
        label_run.font.size = Pt(9)
        label_run.font.bold = True
        label_run.font.color.rgb = RGBColor(*_hex_to_rgb(color))
        count_run = p.add_run(f"  ({count})")
        count_run.font.size = Pt(8)
        count_run.font.color.rgb = RGBColor(120, 120, 120)
    doc.add_paragraph()  # spacer


def _docx_bias_change_card(doc: Document, change: dict) -> None:
    """Add a single colour-coded bias change card as a 1-cell table."""
    color = _BIAS_HEX.get(change["bias_type"], "999999")

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = table.cell(0, 0)
    _set_cell_border_left(cell, color, width=18)
    _hide_cell_borders(cell)
    _set_cell_shading(cell, "FAFAFA")

    # Bias type tag
    p = cell.paragraphs[0]
    tag_run = p.add_run(f"  {_BIAS_LABELS.get(change['bias_type'], change['bias_type']).upper()}  ")
    tag_run.font.size = Pt(7)
    tag_run.font.bold = True
    tag_run.font.color.rgb = RGBColor(255, 255, 255)
    # Simulate tag background via shading on the run
    rpr = tag_run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color.upper())
    rpr.append(shd)

    # Original phrase
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(4)
    p2.paragraph_format.space_after = Pt(1)
    orig_label = p2.add_run("Original: ")
    orig_label.font.size = Pt(9)
    orig_label.font.bold = True
    orig_text = p2.add_run(f'"{change["original_phrase"]}"')
    orig_text.font.size = Pt(9)
    orig_text.font.color.rgb = RGBColor(192, 57, 43)  # red
    orig_text.font.strike = True

    # Replacement phrase
    p3 = cell.add_paragraph()
    p3.paragraph_format.space_before = Pt(1)
    p3.paragraph_format.space_after = Pt(1)
    repl_label = p3.add_run("Replacement: ")
    repl_label.font.size = Pt(9)
    repl_label.font.bold = True
    repl_text = p3.add_run(f'"{change["replacement_phrase"]}"')
    repl_text.font.size = Pt(9)
    repl_text.font.color.rgb = RGBColor(39, 174, 96)  # green
    repl_text.font.bold = True

    # Explanation
    p4 = cell.add_paragraph()
    p4.paragraph_format.space_before = Pt(2)
    p4.paragraph_format.space_after = Pt(2)
    expl_run = p4.add_run(change["explanation"])
    expl_run.font.size = Pt(8)
    expl_run.font.italic = True
    expl_run.font.color.rgb = RGBColor(100, 100, 100)


# ── Main export functions ──

def export_docx(
    text: str,
    output_path: str | Path,
    title: str = "Debiased Report",
    entities_found: list[dict] | None = None,
    changes_summary: str | None = None,
    bias_changes: list[dict] | None = None,
    acronyms_preserved: list[dict] | None = None,
) -> Path:
    """Export debiased report with coloured bias analysis to Word."""
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

    # --- Section 2: Bias Analysis (colour-coded) ---
    if bias_changes:
        doc.add_page_break()
        doc.add_heading(f"Bias Analysis - {len(bias_changes)} Issues Found", level=2)

        intro = doc.add_paragraph(
            "The following biased phrases were identified and corrected. "
            "Each item shows the bias type, original phrasing, neutral replacement, "
            "and an explanation."
        )
        intro.runs[0].font.italic = True
        intro.runs[0].font.size = Pt(9)
        intro.runs[0].font.color.rgb = RGBColor(100, 100, 100)

        # Summary count bar
        _docx_bias_summary_table(doc, bias_changes)

        # Individual change cards
        for change in bias_changes:
            _docx_bias_change_card(doc, change)

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


# ── PDF helpers ──

def _pdf_section_heading(pdf: FPDF, text: str) -> None:
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(30, 30, 60)
    pdf.cell(0, 9, text, new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(30, 30, 60)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(4)


def _pdf_bias_summary_bar(pdf: FPDF, bias_changes: list[dict]) -> None:
    """Draw a row of colour-coded bias type count chips."""
    counts = Counter(c["bias_type"] for c in bias_changes)
    x_start = pdf.l_margin
    for bias_type, count in counts.most_common():
        color = _BIAS_HEX.get(bias_type, "999999")
        r, g, b = _hex_to_rgb(color)
        lr, lg, lb = _hex_to_rgb(_lighten_hex(color, 0.75))
        label = _BIAS_LABELS.get(bias_type, bias_type)
        chip_text = f" {label} ({count}) "

        pdf.set_font("Helvetica", "B", 8)
        chip_w = pdf.get_string_width(chip_text) + 4

        # Check if chip fits on current line
        if pdf.get_x() + chip_w > pdf.w - pdf.r_margin:
            pdf.ln(8)

        # Draw chip background
        pdf.set_fill_color(lr, lg, lb)
        pdf.set_draw_color(r, g, b)
        chip_x = pdf.get_x()
        chip_y = pdf.get_y()
        pdf.rect(chip_x, chip_y, chip_w, 7, style="FD")
        # Coloured left accent
        pdf.set_fill_color(r, g, b)
        pdf.rect(chip_x, chip_y, 2, 7, style="F")

        pdf.set_text_color(r, g, b)
        pdf.set_xy(chip_x, chip_y)
        pdf.cell(chip_w, 7, chip_text)
        pdf.set_x(pdf.get_x() + 3)  # gap between chips

    pdf.ln(10)


def _pdf_bias_change_card(pdf: FPDF, change: dict) -> None:
    """Draw a single colour-coded bias change card in the PDF."""
    color = _BIAS_HEX.get(change["bias_type"], "999999")
    r, g, b = _hex_to_rgb(color)

    card_x = pdf.l_margin
    card_w = pdf.w - pdf.l_margin - pdf.r_margin
    card_start_y = pdf.get_y()

    # Check if we need a new page (estimate ~28mm per card)
    if card_start_y + 28 > pdf.h - pdf.b_margin:
        pdf.add_page()
        card_start_y = pdf.get_y()

    # Card background
    pdf.set_fill_color(250, 250, 250)
    pdf.set_draw_color(230, 230, 230)

    # We'll draw the background after calculating height, so save position
    content_x = card_x + 5  # indent past left accent bar

    # Bias type tag
    pdf.set_xy(content_x, card_start_y + 2)
    tag_label = f"  {_BIAS_LABELS.get(change['bias_type'], change['bias_type']).upper()}  "
    pdf.set_font("Helvetica", "B", 7)
    tag_w = pdf.get_string_width(tag_label) + 2
    pdf.set_fill_color(r, g, b)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(tag_w, 5, tag_label, fill=True)
    pdf.ln(6)

    # Original phrase
    pdf.set_x(content_x)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(18, 5, "Original: ")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(192, 57, 43)
    orig_text = _sanitize_for_pdf(f'"{change["original_phrase"]}"')
    pdf.multi_cell(card_w - 23, 5, orig_text)

    # Replacement phrase
    pdf.set_x(content_x)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(24, 5, "Replacement: ")
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(39, 174, 96)
    repl_text = _sanitize_for_pdf(f'"{change["replacement_phrase"]}"')
    pdf.multi_cell(card_w - 29, 5, repl_text)

    # Explanation
    pdf.set_x(content_x)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(100, 100, 100)
    explanation = _sanitize_for_pdf(change["explanation"])
    pdf.multi_cell(card_w - 10, 4, explanation)

    card_end_y = pdf.get_y() + 2

    # Draw card border and accent bar over the content
    pdf.set_fill_color(250, 250, 250)
    pdf.set_draw_color(230, 230, 230)
    pdf.rect(card_x, card_start_y, card_w, card_end_y - card_start_y, style="D")
    # Left accent bar
    pdf.set_fill_color(r, g, b)
    pdf.rect(card_x, card_start_y, 3, card_end_y - card_start_y, style="F")

    pdf.set_y(card_end_y + 3)


def export_pdf(
    text: str,
    output_path: str | Path,
    title: str = "Debiased Report",
    entities_found: list[dict] | None = None,
    changes_summary: str | None = None,
    bias_changes: list[dict] | None = None,
    acronyms_preserved: list[dict] | None = None,
) -> Path:
    """Export debiased report with coloured bias analysis to PDF."""
    output_path = Path(output_path)
    title = _sanitize_for_pdf(title)
    text = _sanitize_for_pdf(text)
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

    # --- Page 2: Bias Analysis (colour-coded) ---
    if bias_changes:
        pdf.add_page()
        _pdf_section_heading(pdf, f"Bias Analysis - {len(bias_changes)} Issues Found")

        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(
            0, 5,
            "The following biased phrases were identified and corrected. "
            "Each item shows the bias type, original phrasing, neutral replacement, "
            "and an explanation.",
        )
        pdf.ln(4)

        # Summary bar
        _pdf_bias_summary_bar(pdf, bias_changes)

        # Individual change cards
        for change in bias_changes:
            _pdf_bias_change_card(pdf, change)

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
                _truncate(_sanitize_for_pdf(entity["original"]), 30),
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
            pdf.cell(30, 6, _sanitize_for_pdf(acr["text"]), border=1)
            pdf.cell(50, 6, _sanitize_for_pdf(acr["detected_as"]), border=1)
            pdf.cell(100, 6, _truncate(_sanitize_for_pdf(acr["reason"]), 55), border=1)
            pdf.ln()

    pdf.output(str(output_path))
    logger.info("Exported PDF to %s", output_path)
    return output_path


def _sanitize_for_pdf(text: str) -> str:
    """Replace Unicode characters that fpdf2's built-in fonts can't encode."""
    replacements = {
        "\u2018": "'",   # left single quote
        "\u2019": "'",   # right single quote
        "\u201c": '"',   # left double quote
        "\u201d": '"',   # right double quote
        "\u2013": "-",   # en dash
        "\u2014": "-",   # em dash
        "\u2026": "...", # ellipsis
        "\u00a0": " ",   # non-breaking space
        "\u2022": "*",   # bullet
        "\u200b": "",    # zero-width space
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    return text


def _truncate(text: str, max_len: int) -> str:
    return text if len(text) <= max_len else text[: max_len - 3] + "..."


# ── Formatted PDF export (preserve original layout) ──


def _map_to_base14(flags: int) -> str:
    """Map PyMuPDF span flags to the best-matching base14 font name.

    Span flag bits: 1=italic, 2=serif, 3=monospaced, 4=bold.
    """
    is_bold = bool(flags & (1 << 4))
    is_italic = bool(flags & (1 << 1))
    is_serif = bool(flags & (1 << 2))
    is_mono = bool(flags & (1 << 3))

    if is_mono:
        if is_bold and is_italic:
            return "cobi"
        if is_bold:
            return "cobo"
        if is_italic:
            return "coit"
        return "cour"
    if is_serif:
        if is_bold and is_italic:
            return "tibi"
        if is_bold:
            return "tibo"
        if is_italic:
            return "tiit"
        return "tiro"
    if is_bold and is_italic:
        return "hebi"
    if is_bold:
        return "hebo"
    if is_italic:
        return "heit"
    return "helv"


def _get_font_info(
    page_dict: dict, rect: fitz.Rect
) -> tuple[str, float, tuple[float, float, float]]:
    """Find the font name, size, and colour of text nearest to a rect.

    Returns (base14_fontname, fontsize, (r, g, b)) with 0-1 float colours.
    """
    target_y = (rect.y0 + rect.y1) / 2
    target_x = rect.x0
    best = None
    best_dist = float("inf")

    for block in page_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                bbox = span["bbox"]
                dy = abs((bbox[1] + bbox[3]) / 2 - target_y)
                dx = abs(bbox[0] - target_x)
                dist = dy * 2 + dx
                if dist < best_dist:
                    best_dist = dist
                    c = span.get("color", 0)
                    color = (((c >> 16) & 0xFF) / 255, ((c >> 8) & 0xFF) / 255, (c & 0xFF) / 255)
                    best = (_map_to_base14(span.get("flags", 0)), span["size"], color)

    return best or ("helv", 11, (0, 0, 0))


def _export_via_form_fields(
    doc: fitz.Document, changes: list[tuple[str, str]]
) -> bool:
    """Try to apply corrections via PDF form field widgets.

    Returns True if any form fields were found and updated.
    """
    updated = False
    for page in doc:
        for widget in page.widgets():
            value = widget.field_value
            if not value:
                continue
            modified = value
            changed = False
            for original, replacement in changes:
                if original in modified:
                    modified = modified.replace(original, replacement)
                    changed = True
            if changed:
                widget.field_value = modified
                widget.update()
                updated = True
                logger.debug("Updated field '%s' on page %d", widget.field_name, page.number)
    return updated


def _export_via_text_replacement(
    doc: fitz.Document, changes: list[tuple[str, str]]
) -> None:
    """Apply corrections by surgically replacing individual phrases in the PDF.

    For each biased phrase: finds its exact position with search_for(),
    redacts just that small rect (preserving all surrounding content),
    then inserts the replacement text at the same baseline position.
    """
    for page in doc:
        page_dict = page.get_text("dict")
        pending: list[tuple[fitz.Rect, str, str, float, tuple]] = []

        for original, replacement in changes:
            rects = page.search_for(original)
            if not rects:
                continue

            for rect in rects:
                fontname, fontsize, color = _get_font_info(page_dict, rect)
                page.add_redact_annot(rect)
                pending.append((rect, replacement, fontname, fontsize, color))

            logger.debug(
                "Found '%s' → '%s' (%d locations on page %d)",
                original, replacement, len(rects), page.number,
            )

        if not pending:
            continue

        # Remove only the text — preserve all graphics (form borders, lines)
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_NONE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
        )

        # Insert replacement text at the original baseline position
        for rect, replacement, fontname, fontsize, color in pending:
            # Shrink only if replacement is wider than the available space
            text_width = fitz.get_text_length(replacement, fontname=fontname, fontsize=fontsize)
            if text_width > rect.width:
                # Extend rect rightward (up to page margin) before shrinking
                max_width = page.rect.width - 36 - rect.x0
                rect = fitz.Rect(rect.x0, rect.y0, rect.x0 + max(rect.width, max_width), rect.y1)
                text_width = fitz.get_text_length(replacement, fontname=fontname, fontsize=fontsize)
                if text_width > rect.width:
                    fontsize = fontsize * (rect.width / text_width) * 0.98

            baseline_y = rect.y0 + fontsize * 0.88
            page.insert_text(
                (rect.x0, baseline_y), replacement,
                fontname=fontname, fontsize=fontsize, color=color,
            )


def _export_formatted_native(
    doc: fitz.Document, bias_changes: list[dict], entity_mapping: dict[str, str]
) -> None:
    """Apply bias corrections to a native PDF, preserving form layout.

    Strategy:
    1. If the PDF has fillable form fields → update field values directly.
    2. Otherwise → surgically replace individual phrases in-place, using
       small per-phrase redactions that don't disturb surrounding content.
    """
    changes = [
        (
            unmask_phrase(c["original_phrase"], entity_mapping),
            unmask_phrase(c["replacement_phrase"], entity_mapping),
        )
        for c in bias_changes
    ]

    if not _export_via_form_fields(doc, changes):
        logger.info("No form fields found; falling back to text replacement")
        _export_via_text_replacement(doc, changes)


def _export_formatted_scanned(
    doc: fitz.Document,
    original_pdf_path: str | Path,
    bias_changes: list[dict],
    entity_mapping: dict[str, str],
) -> None:
    """Apply bias corrections on a scanned PDF using OCR bounding boxes.

    Converts each page to an image, runs OCR to get word-level bounding boxes
    grouped into blocks, applies bias corrections per block, then whites out
    the old text and overlays the corrected text.
    """
    import pytesseract
    from pdf2image import convert_from_path
    from pytesseract import Output

    changes = [
        (
            unmask_phrase(c["original_phrase"], entity_mapping),
            unmask_phrase(c["replacement_phrase"], entity_mapping),
        )
        for c in bias_changes
    ]

    dpi = 300
    scale = 72.0 / dpi
    images = convert_from_path(str(original_pdf_path), dpi=dpi)

    for page_idx, image in enumerate(images):
        if page_idx >= len(doc):
            break
        page = doc[page_idx]
        ocr_data = pytesseract.image_to_data(image, output_type=Output.DICT)

        # Group OCR words into blocks (Tesseract provides block_num)
        blocks: dict[int, list[int]] = {}
        for i in range(len(ocr_data["text"])):
            if not ocr_data["text"][i].strip():
                continue
            blocks.setdefault(ocr_data["block_num"][i], []).append(i)

        for bnum, indices in blocks.items():
            block_text = " ".join(ocr_data["text"][i] for i in indices)

            modified = block_text
            has_changes = False
            for original, replacement in changes:
                if original in modified:
                    modified = modified.replace(original, replacement)
                    has_changes = True

            if not has_changes:
                continue

            # Compute encompassing rectangle (pixel → PDF points)
            px_left = min(ocr_data["left"][i] for i in indices)
            px_top = min(ocr_data["top"][i] for i in indices)
            px_right = max(ocr_data["left"][i] + ocr_data["width"][i] for i in indices)
            px_bottom = max(ocr_data["top"][i] + ocr_data["height"][i] for i in indices)
            rect = fitz.Rect(
                px_left * scale, px_top * scale,
                px_right * scale, px_bottom * scale,
            )

            avg_height = sum(ocr_data["height"][i] for i in indices) / len(indices)
            fontsize = avg_height * scale * 0.85

            page.draw_rect(rect, color=None, fill=(1, 1, 1))
            page.insert_textbox(
                rect, modified,
                fontsize=fontsize, fontname="helv", color=(0, 0, 0),
            )
            logger.debug("Replaced OCR block %d on page %d", bnum, page_idx)


def _add_searchable_text_layer(doc: fitz.Document) -> None:
    """Add an invisible OCR text layer to pages that lack extractable text.

    Renders each page to an image, runs Tesseract, and inserts every detected
    word as invisible text (render_mode=3).  The text is in the content stream
    so Adobe Acrobat (and other readers) can search / select it, but it doesn't
    change the visual appearance.

    Pages that already have substantial extractable text are skipped to avoid
    duplicate search hits.
    """
    import pytesseract
    from PIL import Image
    from pytesseract import Output

    dpi = 300
    scale = 72.0 / dpi

    for page in doc:
        # Skip pages that already have a usable text layer
        if len(page.get_text().strip()) > 50:
            continue

        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ocr_data = pytesseract.image_to_data(img, output_type=Output.DICT)

        for i in range(len(ocr_data["text"])):
            word = ocr_data["text"][i].strip()
            if not word:
                continue

            x = ocr_data["left"][i] * scale
            y = ocr_data["top"][i] * scale
            h = ocr_data["height"][i] * scale
            fontsize = max(h * 0.85, 1)

            baseline_y = y + fontsize * 0.88
            page.insert_text(
                (x, baseline_y), word,
                fontname="helv", fontsize=fontsize,
                render_mode=3,  # invisible but searchable
            )

        logger.debug("Added searchable text layer to page %d", page.number)


def export_formatted_pdf(
    original_pdf_path: str | Path,
    bias_changes: list[dict],
    entity_mapping: dict[str, str],
    output_path: str | Path,
    is_scanned: bool = False,
) -> Path:
    """Export a formatted PDF with bias corrections applied in-place.

    For native PDFs with form fields: updates field values directly.
    For native PDFs without form fields: surgical per-phrase text replacement.
    For scanned PDFs: OCR bounding boxes to white-out and overlay text.

    After corrections, adds an invisible OCR text layer to any page that
    lacks extractable text, making the output searchable in Adobe Acrobat.
    """
    output_path = Path(output_path)
    doc = fitz.open(str(original_pdf_path))

    if not bias_changes:
        logger.info("No bias changes to apply; saving original PDF as-is.")
        doc.save(str(output_path))
        doc.close()
        return output_path

    if is_scanned:
        _export_formatted_scanned(doc, original_pdf_path, bias_changes, entity_mapping)
    else:
        _export_formatted_native(doc, bias_changes, entity_mapping)

    # Make non-searchable pages searchable via invisible OCR text layer
    _add_searchable_text_layer(doc)

    doc.save(str(output_path), garbage=3, deflate=True)
    doc.close()
    logger.info("Exported formatted PDF to %s", output_path)
    return output_path
