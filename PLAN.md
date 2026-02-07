# AI Report Assist — Implementation Plan

## Context

Law enforcement reports can contain implicit biases that affect fairness and justice outcomes. This tool processes PDF reports (native or scanned) through a 6-step pipeline:

1. Convert PDF to text
2. Identify sensitive information (LOCAL)
3. Mask sensitive information (LOCAL)
4. Send masked text to AI for debiasing (cloud — safe because PII is masked)
5. Unmask the debiased content (LOCAL)
6. Review and export

**All PII handling stays on the local machine.** Only anonymized/masked text is sent to the cloud API.

---

## Environment

- **Conda environment**: `ai-report-assist` (Python 3.11)
- **IDE**: Cursor — select interpreter via `Cmd+Shift+P → Python: Select Interpreter → ai-report-assist`
- **System deps installed via conda**: `tesseract`, `poppler`

---

## Technology Stack

| Layer | Technology | Why |
|-------|-----------|-----|
| Backend | FastAPI + Uvicorn | Async, fast, auto-generated API docs |
| Frontend | Jinja2 templates + HTMX | Simple, no JS build step, server-rendered |
| PDF (native) | PyMuPDF | Fast, reliable text extraction |
| PDF (scanned/OCR) | pdf2image + pytesseract | Industry-standard OCR |
| PII detection | Presidio Analyzer + spaCy `en_core_web_lg` | Best local PII detection, fully offline |
| Custom entities | Presidio PatternRecognizer | Regex + context words for law enforcement entities |
| Masking | Presidio Anonymizer | Token replacement with reversible mapping |
| Unmasking | Presidio Deanonymizer | Restore original values from mapping |
| AI debiasing | OpenAI API (GPT-4o) | High-quality bias analysis on masked text |
| Export (DOCX) | python-docx | Native Word format |
| Export (PDF) | fpdf2 | Simple, Unicode-friendly PDF generation |

---

## Project Structure

```
ai-report-assist/
├── PLAN.md                       # This file
├── environment.yml               # Conda environment definition
├── .env.example                  # Template for API keys
├── .gitignore
│
├── app/
│   ├── __init__.py
│   ├── main.py                   # FastAPI app entry point + routes
│   ├── config.py                 # Settings (env vars, paths, model config)
│   │
│   ├── pipeline/
│   │   ├── __init__.py
│   │   ├── extractor.py          # Step 1: PDF → text
│   │   ├── analyzer.py           # Step 2: Detect PII entities
│   │   ├── masker.py             # Step 3: Mask PII + generate mapping
│   │   ├── debiaser.py           # Step 4: Send to OpenAI for debiasing
│   │   ├── unmasker.py           # Step 5: Restore original PII
│   │   └── exporter.py           # Step 6: Export to PDF / DOCX
│   │
│   ├── recognizers/
│   │   ├── __init__.py
│   │   └── law_enforcement.py    # Custom Presidio recognizers
│   │
│   ├── templates/
│   │   ├── base.html             # Shared layout
│   │   ├── upload.html           # File upload page
│   │   ├── review.html           # Side-by-side: original vs debiased
│   │   └── export.html           # Format selection + download
│   │
│   └── static/
│       └── style.css             # Minimal CSS
│
├── uploads/                      # Temporary uploaded PDFs
│
└── tests/
    ├── test_extractor.py
    ├── test_analyzer.py
    ├── test_masker.py
    └── test_pipeline.py
```

---

## Pipeline Details

### Step 1: PDF Text Extraction (`extractor.py`)

- Auto-detect native vs scanned PDF
  - Try `pymupdf` text extraction first
  - If result is empty/near-empty → fall back to OCR
- **Native**: `pymupdf` → `page.get_text()` per page
- **Scanned**: `pdf2image.convert_from_path()` → `pytesseract.image_to_string()` per image
- Output: `{ pages: [{ page_num, text }], is_scanned: bool }`

### Step 2: PII Detection (`analyzer.py` + `law_enforcement.py`)

**Built-in Presidio entities:**
- PERSON, LOCATION, PHONE_NUMBER, EMAIL_ADDRESS, DATE_TIME, US_SSN, CREDIT_CARD, IP_ADDRESS, US_DRIVER_LICENSE

**Custom law enforcement recognizers (PatternRecognizer):**

| Entity | Regex Pattern | Context Words |
|--------|--------------|---------------|
| `BADGE_NUMBER` | `[A-Z]{0,3}\d{4,7}` | badge, officer id, shield |
| `CASE_NUMBER` | `\d{2,4}[-/]\d{2,6}[-/]?\d{0,6}` | case, docket, CR, file no |
| `EVIDENCE_ID` | `EV[-]?\d{4,8}` | evidence, exhibit, item |
| `LICENSE_PLATE` | `[A-Z0-9]{1,4}[\s-]?[A-Z0-9]{2,5}` | plate, license, vehicle, registration, tag |
| `WEAPON_SERIAL` | `[A-Z]{0,3}\d{4,10}[A-Z]{0,2}` | serial, weapon, firearm, gun, pistol, rifle |

Context words increase confidence scores and reduce false positives.

### Step 3: Masking (`masker.py`)

- Token replacement: `John Smith` → `[PERSON_1]`, `123 Main St` → `[ADDRESS_1]`
- **Consistent**: same real entity always maps to the same token throughout the document
- **Reversible**: store bidirectional mapping dictionary
- Output: `{ masked_text, entity_mapping, entities_found }`

### Step 4: AI Debiasing (`debiaser.py`)

- Send **only masked text** to OpenAI GPT-4o
- Prompt instructs the model to:
  1. Remove language suggesting demographic profiling
  2. Remove stereotypical assumptions
  3. Replace subjective characterizations with neutral factual language
  4. Remove inflammatory or prejudicial word choices
  5. Preserve all factual content and entity tokens exactly as-is
- Return: `{ debiased_text, changes_summary }`

### Step 5: Unmasking (`unmasker.py`)

- Replace all tokens with original values using the mapping
- Validate that no tokens remain unresolved
- Output: `{ final_text, unresolved_tokens }`

### Step 6: Export (`exporter.py`)

- **DOCX**: `python-docx` — document with title, date, body paragraphs
- **PDF**: `fpdf2` — formatted report with metadata
- Both include: report title, processing date, page numbers

---

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/` | Upload page |
| `POST` | `/upload` | Accept PDF, extract text, return job ID |
| `POST` | `/process/{job_id}` | Run mask → debias → unmask pipeline |
| `GET` | `/review/{job_id}` | Side-by-side review page |
| `POST` | `/export/{job_id}` | Generate and download PDF or DOCX |

---

## Frontend Pages

1. **Upload** — Drag-and-drop PDF upload with file type validation
2. **Processing** — Progress indicator while pipeline runs
3. **Review** — Side-by-side view of original text vs debiased text, with changes highlighted
4. **Export** — Choose format (PDF / DOCX), download button

---

## Implementation Order

| Phase | What | Files |
|-------|------|-------|
| 1 | Environment + project skeleton | `environment.yml`, directory structure, `main.py` |
| 2 | PDF extraction | `extractor.py` |
| 3 | PII detection + custom recognizers | `analyzer.py`, `law_enforcement.py` |
| 4 | Masking with mapping | `masker.py` |
| 5 | OpenAI debiasing | `debiaser.py` |
| 6 | Unmasking | `unmasker.py` |
| 7 | Export | `exporter.py` |
| 8 | Frontend + routes | `main.py`, templates, static |
| 9 | Testing | `tests/` |

---

## Dependencies (`environment.yml`)

```yaml
name: ai-report-assist
channels:
  - conda-forge
  - defaults
dependencies:
  - python=3.11
  - tesseract
  - poppler
  - pip
  - pip:
      - fastapi
      - uvicorn[standard]
      - python-multipart
      - jinja2
      - python-dotenv
      - pymupdf
      - pdf2image
      - pytesseract
      - presidio-analyzer
      - presidio-anonymizer
      - spacy
      - openai
      - python-docx
      - fpdf2
      - httpx
      - pytest
```

Post-install: `python -m spacy download en_core_web_lg`

---

## Verification Checklist

- [ ] Conda env activates and `python` points to env interpreter
- [ ] Cursor IDE uses the conda env interpreter
- [ ] Upload a native PDF → text extracted correctly
- [ ] Upload a scanned PDF → OCR produces readable text
- [ ] PII entities detected (names, addresses, SSN, phone, etc.)
- [ ] Custom entities detected (badge numbers, case numbers, plates, etc.)
- [ ] Masked text contains zero real PII
- [ ] OpenAI receives only masked text
- [ ] Debiased text preserves facts, removes biased language
- [ ] Unmasked text has all original identifiers restored
- [ ] Export to PDF opens correctly
- [ ] Export to DOCX opens correctly
