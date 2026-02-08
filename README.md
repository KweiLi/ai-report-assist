# AI Report Assist

A privacy-first law enforcement report debiasing tool. Upload a police report PDF, and the pipeline automatically detects and masks PII locally, sends only masked text to AI for bias correction, then restores PII and exports a clean report preserving the original document layout.

## Tech Stack

| Layer | Technology |
|-------|-----------|
| **Backend** | Python 3.11, FastAPI, Uvicorn |
| **Frontend** | Jinja2 templates, HTMX, vanilla CSS |
| **PII Detection** | Presidio Analyzer + spaCy (`en_core_web_lg`) — runs entirely local |
| **Custom Recognizers** | Badge numbers, case numbers, evidence IDs, license plates, weapon serials |
| **AI Debiasing** | OpenAI GPT-4o (receives only masked text) |
| **PDF Processing** | PyMuPDF (fitz), pdf2image, Tesseract OCR, pytesseract |
| **Export** | PyMuPDF (formatted PDF), fpdf2 (analysis PDF), python-docx (DOCX) |
| **Environment** | Conda (tesseract + poppler as system deps) |

## Privacy Architecture

```
 User's Machine (LOCAL)                          Cloud
 ┌─────────────────────────────────────┐        ┌──────────┐
 │  PDF ──► Extract ──► Analyze PII    │        │          │
 │                        │            │        │  OpenAI  │
 │              Mask (replace PII      │        │  GPT-4o  │
 │              with tokens like       │ ──────►│          │
 │              [PERSON_1], [ADDR_1])  │ masked │          │
 │                                     │  text  │  Debias  │
 │              Unmask (restore PII    │◄────── │          │
 │              from token mapping)    │ result │          │
 │                        │            │        └──────────┘
 │              Export ──► PDF / DOCX  │
 └─────────────────────────────────────┘
```

- PII detection and masking happen **entirely on-device** using Presidio + spaCy
- Only **masked text** (with tokens like `[PERSON_1]`, `[LOCATION_2]`) is sent to OpenAI
- The token-to-original mapping never leaves the local machine
- Unmasking restores real names/addresses after debiasing

## Pipeline Steps

1. **Upload** — Accept PDF (native or scanned)
2. **Extract** — PyMuPDF for native PDFs; Tesseract OCR for scanned documents
3. **Analyze PII** — Presidio with spaCy NLP + custom law enforcement recognizers; preserves ~150 standard LE acronyms (BOLO, DUI, SWAT, etc.)
4. **Mask** — Replace each PII entity with a typed token (`[PERSON_1]`, `[BADGE_NUM_1]`, etc.)
5. **Debias** — Send masked text to GPT-4o with structured prompt; returns bias changes with categories (racial/ethnic, gender, socioeconomic, stereotyping, inflammatory, confirmation, subjective)
6. **Unmask** — Restore all tokens to original values
7. **Review** — Side-by-side comparison with color-coded bias highlights
8. **Export** — Three formats:
   - **Formatted PDF** — Bias corrections applied in-place on the original PDF layout (form fields, headers, borders preserved)
   - **Analysis PDF** — Full debiased text + color-coded bias analysis cards + masking details
   - **Analysis DOCX** — Same as analysis PDF in editable Word format

## Project Structure

```
app/
├── main.py                    # FastAPI routes and job management
├── config.py                  # Environment config (API keys, thresholds)
├── templates/
│   ├── base.html              # Base layout
│   ├── upload.html            # File upload page
│   └── review.html            # Pipeline review + export page
├── static/
│   └── style.css              # UI styles
├── pipeline/
│   ├── extractor.py           # PDF text extraction (native + OCR)
│   ├── analyzer.py            # PII detection via Presidio
│   ├── masker.py              # PII masking (token replacement)
│   ├── debiaser.py            # OpenAI GPT-4o bias correction
│   ├── unmasker.py            # Token-to-original restoration
│   └── exporter.py            # PDF/DOCX export (formatted + analysis)
└── recognizers/
    └── law_enforcement.py     # Custom Presidio recognizers + acronym list
environment.yml                # Conda environment spec
```

## Setup

### Prerequisites

- [Conda](https://docs.conda.io/) (Miniconda or Anaconda)
- An OpenAI API key

### Install

```bash
# Create conda environment (installs tesseract + poppler + all Python deps)
conda env create -f environment.yml
conda activate ai-report-assist

# Download spaCy language model
python -m spacy download en_core_web_lg

# Configure environment
cp .env.example .env
# Edit .env and add your OPENAI_API_KEY
```

### Run

```bash
uvicorn app.main:app --reload
```

Open http://localhost:8000 in your browser.

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `OPENAI_API_KEY` | (required) | OpenAI API key for GPT-4o debiasing |
| `OPENAI_MODEL` | `gpt-4o` | OpenAI model to use |
| `PII_CONFIDENCE_THRESHOLD` | `0.55` | Minimum confidence for PII detection (0.0-1.0) |
