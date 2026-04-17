# OMR Grader App — Session Summary
*Paste this at the start of a new chat to continue development.*

---

## 🤖 Instruction for Claude
**Before making any code changes, always ask:**
> "I can do this the token-efficient way (surgical `str_replace` on affected lines only) or follow your specific instructions. Which do you prefer?"

**Additional token-saving rules:**
- Use `str_replace` for targeted edits — never rewrite full files unless the change touches >50% of the file
- Batch cosmetic changes together and preview once
- Skip visual previews for logic-only changes (scanner, grader, exporter)
- When confirming before building, keep the analysis concise — avoid re-explaining things already established
- If the user says "go ahead" or similar, proceed without further confirmation

---

## Project Overview
Desktop GUI (Tkinter) Python app for grading OMR exam sheets.  
Built for **Fundación Educa México A.C.** — girlfriend's classroom use.

**Workflow:**
1. Generate printable OMR answer sheets as PDF
2. Scan completed sheets (scanned PDF input)
3. Grade Section 1 (MC) against answer key
4. Record Section 2 (self-knowledge scale 1–5)
5. Export results to Excel with charts

---

## File Structure
All files live in `omr_app/`:

| File | Purpose |
|------|---------|
| `layout.py` | **Single source of truth** for ALL coordinates — both `sheet_generator.py` and `omr_scanner.py` import from here. Never hardcode positions elsewhere. |
| `sheet_generator.py` | Generates OMR PDFs via ReportLab |
| `omr_scanner.py` | Detects bubbles via OpenCV + pdf2image |
| `omr_grader.py` | Scores MC answers, passes through metadata |
| `omr_exporter.py` | Builds Excel output (openpyxl) |
| `main.py` | Tkinter GUI |
| `omr_config.json` | Saved settings |
| `logo.png` | Fundación Educa México A.C. logo (RGBA, 3194×959px, aspect 3.331) |

---

## Current Sheet Layout (Letter, 612×792pt)

### Header
- **Corner markers:** 18×18pt black squares, 45pt from each page edge (~16mm, printer-safe)
- **Logo:** Top-left, 1.8in wide × 38.9pt tall, top edge flush with marker top (`LOGO_Y = PAGE_H - 45 - 9 - 38.9`)
- **Exam title:** Centered on page (`PAGE_W/2`), same vertical center as logo, 15pt Helvetica-Bold, navy `#1B2A6B`

### Field Row 1 (text, 12pt)
`Nombre: ___`  /  `Fecha: ___`  /  `Escuela: ___`

### Bubble Zone (below Field Row 1)
Grado and Grupo on a single row, centered vertically between the two Folio rows.
Folio sits to the right with two stacked bubble rows.

```
                                    1  2  3  4  5  6  7  8  9   ← digit labels
Grado  1  2  3    Grupo  A–F   Folio  N1: ○  ○  ○  ○  ○  ○  ○  ○  ○
       ○  ○  ○           ○×6          N2: ○  ○  ○  ○  ○  ○  ○  ○  ○
```

- `FOLIO_ROW1_Y = FIELDS_Y1 - 30`   (Número 1, top row)
- `FOLIO_ROW2_Y = FIELDS_Y1 - 50`   (Número 2, bottom row)
- `BUBBLE_ROW_Y = (FOLIO_ROW1_Y + FOLIO_ROW2_Y) / 2`  → `FIELDS_Y1 - 40`  (Grado/Grupo center)
- `SEPARATOR_Y  = FOLIO_ROW2_Y - BUBBLE_R - 8`  → `FIELDS_Y1 - 64.5`

### Section 1 — Opción Múltiple
- Configurable: 5–40 questions, 2 columns
- Labels: A B C D above bubbles; range labels in italic gray
- `SEC1_FIRST_Q_Y = SEC1_TITLE_Y - 42`

### Section 2 — Autoconocimiento
- Fixed: **10 questions**, 2 columns of 5, scale 1–5
- Derived from `sec2_title_y(num_mc_questions)`

### Footer
`"Use lápiz o bolígrafo negro. Rellene completamente el círculo."` — 9pt gray, centered

---

## Key Layout Constants (`layout.py`)

```python
BUBBLE_R      = 6.5    # bubble radius (pt)
BUBBLE_SP_X   = 22     # horizontal spacing between bubble centers
BUBBLE_SP_Y   = 20     # vertical spacing between question rows
Q_NUM_OFFSET  = 20     # question number right-edge x from col start
BUBBLE_OFFSET = 32     # first bubble center x from col start
COL_W         = 2.1 * inch
MARKER_SIZE   = 18
MARKER_OFFSET = 45
SK_Q          = 10     # fixed number of Autoconocimiento questions
FONT_SIZE     = 12     # (in sheet_generator.py)
```

### Grado / Grupo X positions
```python
GRADO_LABEL_X = MARGIN_L
GRADO_B1_X    = GRADO_LABEL_X + 48
GRADO_VALUES  = ["1", "2", "3"]
grado_bubble_x(idx) = GRADO_B1_X + idx * BUBBLE_SP_X

GRUPO_LABEL_X = GRADO_B1_X + 2*BUBBLE_SP_X + BUBBLE_R + 22
GRUPO_B1_X    = GRUPO_LABEL_X + 48
GRUPO_VALUES  = ["A", "B", "C", "D", "E", "F"]
grupo_bubble_x(idx) = GRUPO_B1_X + idx * BUBBLE_SP_X
```

### Folio X/Y positions
```python
FOLIO_ROW1_Y  = FIELDS_Y1 - 30          # Número 1 bubble center Y
FOLIO_ROW2_Y  = FIELDS_Y1 - 50          # Número 2 bubble center Y
FOLIO_LABEL_X = GRUPO_B1_X + (len(GRUPO_VALUES)-1)*BUBBLE_SP_X + BUBBLE_R + 12
FOLIO_B1_X    = FOLIO_LABEL_X + 52      # first bubble center x
FOLIO_SP_X    = 19                       # tighter spacing for 9 bubbles
FOLIO_VALUES  = ["1","2","3","4","5","6","7","8","9"]
folio_bubble_x(idx) = FOLIO_B1_X + idx * FOLIO_SP_X
```

---

## Data Flow

### ScanResult (omr_scanner.py)
```python
@dataclass
class ScanResult:
    page_num:   int
    folio:      str             # scanned 2-digit string e.g. '37', '?5', '??'
    grado:      Optional[str]   # '1'/'2'/'3' / '?' / None
    grupo:      Optional[str]   # 'A'–'F' / '?' / None
    mc_answers: list            # 'A'/'B'/'C'/'D' or None per question
    sk_answers: list            # 1–5 or None per question (10 questions)
    confidence: float
    error:      Optional[str]
```

### GradeResult (omr_grader.py)
Same fields as ScanResult plus:
```python
    mc_correct: list    # bool per question
    score:      int
    total:      int
    percentage: float
    sk_average: Optional[float]
```

### scan_pdf() signature
```python
scan_pdf(pdf_path, num_mc_questions, progress_callback=None)
# folio_start removed — folio is now read from the sheet bubbles
```

---

## Scanner Architecture

- PDF → pdf2image (200 DPI) → numpy arrays
- Corner markers → perspective warp to 850×1100px (`WARP_W/H`)
- Coordinate conversion: `_pt_to_px(x_pt, y_pt)` — handles ReportLab bottom-up → image top-down
- `SCAN_R = 9` (detection radius in warped px)
- **MC/SK answers:** `_best_in_row()` — picks most-filled bubble above threshold 0.42
- **Grado/Grupo/Folio:** `_best_with_ambiguity()` — returns `'?'` if multiple filled, `None` if none filled
- Folio scanning: two rows scanned independently → combined as `f"{d1 or '?'}{d2 or '?'}"`

---

## Excel Output (omr_exporter.py) — 4 Sheets

Column order everywhere: **Folio → Grado → Grupo → ...**  *(Género removed)*

1. **Resumen** — 10 columns: Folio, Grado, Grupo, Puntaje, Total, Porcentaje(%), Calificación, Prom.Autoconoc., Confianza(%), Estado
2. **Detalle Preguntas** — per-question answer grid (cols 1–4 metadata, col 5+ answers), green=correct/red=wrong
3. **Autoconocimiento** — 14 columns: Folio, Grado, Grupo, SK1–SK10, Promedio (merge A1:N1)
4. **Gráficas** — bar chart (scores) + line chart (SK averages)

---

## GUI (main.py)

Settings panel inputs:
- Exam name (text entry)
- Nº preguntas Sec.1 (Spinbox, 5–40)

*(Folio inicial spinbox removed — folio now comes from the scanned sheet)*

Buttons: `Generar Hoja OMR` / `Calificar PDF`

Config saved to `omr_config.json`: `exam_name`, `num_mc_questions`, `answer_key`, `last_output_dir`

---

## Fit Verification
- **40q + 10 SK worst case:** last SK bubble Y ≈ 38pt, footer at 20pt — very tight but fits
- **38q + 10 SK:** last SK bubble Y ≈ 58pt — comfortable ✅
- **Folio right edge:** ~550pt, right margin at ~568pt — 18pt clearance ✅
- **Digit labels vs text row:** 16.5pt gap between text baseline and digit label baseline ✅
- **Corner markers at 45pt** from edge (~16mm) — safe for most printers ✅

---

## Known Design Decisions
- **Folio is student-filled** (two bubble rows N1, N2, digits 1–9 each) — used to match against a database for name and gender lookup
- **Género removed** from sheet and all outputs — gender retrieved via folio→DB lookup instead
- Written fields (Nombre, Fecha, Escuela) kept as handwritten fallback — not scanned
- Grado/Grupo/Folio all use ambiguity detection — two bubbles filled → `'?'`
- Logo top edge flush with top-left corner marker — nothing outside marker boundary
- `layout.py` is the single coordinate source — scanner never has hardcoded positions
- `FOLIO_SP_X = 19` (slightly tighter than standard 22) to fit 9 bubbles in the right margin
