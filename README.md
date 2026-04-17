# Calificador de Exámenes OMR

Desktop app for grading optical mark recognition (OMR) exam sheets, built for **Fundación Educa México A.C. / Amgen Biotech Experience (ABE)**.

## What it does

1. **Generate** — Prints a ready-to-use OMR answer sheet as a PDF
2. **Scan** — Reads a batch of scanned sheets (PDF, ~35–45 pages) using OpenCV
3. **Grade** — Scores Section 1 (multiple choice) against a configurable answer key
4. **Record** — Captures Section 2 (self-knowledge scale 1–5)
5. **Export** — Produces an Excel workbook with scores, charts, and answer detail

## Requirements

- Python 3.9+
- [Poppler](https://poppler.freedesktop.org/) (for PDF → image conversion)

Install Python dependencies:

```bash
pip install -r omr_app/requirements.txt
```

On macOS, install Poppler via Homebrew:

```bash
brew install poppler
```

## Running

```bash
python omr_app/main.py
```

## Sheet format

- **Letter** paper (8.5 × 11 in)
- Section 1: 5–40 multiple-choice questions (A–D), 2 columns
- Section 2: 10 self-knowledge questions (scale 1–5), 2 columns
- Folio: two rows of bubbles (digits 1–9) for student ID lookup
- Grado (1–3) and Grupo (A–F) bubble fields

## Student list

Load an Excel file with two columns — `folio` and `nombre` — to automatically populate student names in the results. Unknown folios are highlighted in yellow in the exported Excel.

## License

Internal use — Fundación Educa México A.C.
