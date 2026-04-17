"""
layout.py — Single source of truth for all OMR sheet coordinates.
Both sheet_generator.py and omr_scanner.py import from here.
All values in ReportLab points (pt). Origin: bottom-left. Letter: 612x792pt.
"""
from reportlab.lib.units import inch

# ── Page ──────────────────────────────────────────────────────────────────────
PAGE_W = 612.0
PAGE_H = 792.0

# ── Margins ───────────────────────────────────────────────────────────────────
MARGIN_L = 0.6 * inch
MARGIN_R = 0.6 * inch

# ── Bubble geometry ───────────────────────────────────────────────────────────
BUBBLE_R      = 6.5
BUBBLE_SP_X   = 22
BUBBLE_SP_Y   = 20
Q_NUM_OFFSET  = 20    # question number right-edge x offset from col start
BUBBLE_OFFSET = 32    # first bubble center x offset from col start

# ── Column layout ─────────────────────────────────────────────────────────────
COL_W = 2.1 * inch

# ── Corner markers ────────────────────────────────────────────────────────────
MARKER_SIZE   = 18
MARKER_OFFSET = 45
MARKERS = [
    (MARKER_OFFSET,           MARKER_OFFSET),
    (PAGE_W - MARKER_OFFSET,  MARKER_OFFSET),
    (MARKER_OFFSET,           PAGE_H - MARKER_OFFSET),
    (PAGE_W - MARKER_OFFSET,  PAGE_H - MARKER_OFFSET),
]

# ── Logo (left) — Fundación Educa México ──────────────────────────────────────
# logo.png native size: 3194×959px → aspect 3.331:1
LOGO_W = 1.1 * inch                    # reduced so Nombre field has room
LOGO_H = round(LOGO_W / 3.331, 1)     # ≈ 23.8pt, preserves aspect ratio
LOGO_X = MARGIN_L
LOGO_Y = PAGE_H - MARKER_OFFSET - MARKER_SIZE/2 - LOGO_H  # top flush with marker top

# ── Logo (right) — ABE ────────────────────────────────────────────────────────
# ABE_LOGO_W is computed dynamically in sheet_generator from the image aspect ratio
ABE_LOGO_H = LOGO_H                    # same height as left logo

# ── Exam title — centered on page at logo vertical center ─────────────────────
TITLE_X = PAGE_W / 2
TITLE_Y = LOGO_Y + LOGO_H * 0.45

# ── Field row 1: Nombre, Fecha, Escuela (text baselines) ──────────────────────
FIELDS_Y1 = LOGO_Y - 30   # extra gap below logos (~15pt font + padding)

# ── Bubble zone: Grado/Grupo on one row; Folio on two rows (right side) ───────
# Folio rows (Número 1 and Número 2) define the zone height
FOLIO_ROW1_Y = FIELDS_Y1 - 30   # Número 1 bubble center Y (top row)
FOLIO_ROW2_Y = FIELDS_Y1 - 50   # Número 2 bubble center Y (bottom row)

# Grado/Grupo: single row centered between the two Folio rows
BUBBLE_ROW_Y = (FOLIO_ROW1_Y + FOLIO_ROW2_Y) / 2   # FIELDS_Y1 - 34

# Grado bubbles: 3 options (1, 2, 3)
GRADO_LABEL_X = MARGIN_L
GRADO_B1_X    = GRADO_LABEL_X + 48   # first bubble center (gap after label)
GRADO_VALUES  = ["1", "2", "3"]

def grado_bubble_x(idx):
    return GRADO_B1_X + idx * BUBBLE_SP_X

# Grupo bubbles: 6 options (A–F)
GRUPO_LABEL_X = GRADO_B1_X + 2 * BUBBLE_SP_X + BUBBLE_R + 22  # after grado + gap
GRUPO_B1_X    = GRUPO_LABEL_X + 48
GRUPO_VALUES  = ["A", "B", "C", "D", "E", "F"]

def grupo_bubble_x(idx):
    return GRUPO_B1_X + idx * BUBBLE_SP_X

# Folio bubbles: 2 rows × 9 options (1–9)
FOLIO_LABEL_X = GRUPO_B1_X + (len(GRUPO_VALUES) - 1) * BUBBLE_SP_X + BUBBLE_R + 12
FOLIO_B1_X    = FOLIO_LABEL_X + 52   # first bubble center x
FOLIO_SP_X    = 19                    # tighter spacing for 9 bubbles
FOLIO_VALUES  = ["1", "2", "3", "4", "5", "6", "7", "8", "9"]

def folio_bubble_x(idx):
    return FOLIO_B1_X + idx * FOLIO_SP_X

# ── Separator below bubble zone ───────────────────────────────────────────────
SEPARATOR_Y = FOLIO_ROW2_Y - BUBBLE_R - 8

# ── Section 1 ─────────────────────────────────────────────────────────────────
SEC1_TITLE_Y   = SEPARATOR_Y - 18
SEC1_RANGE_Y   = SEC1_TITLE_Y - 14
SEC1_HDR_Y     = SEC1_TITLE_Y - 28
SEC1_FIRST_Q_Y = SEC1_TITLE_Y - 42

def sec1_bubble(q_idx, choice_idx, num_mc_questions):
    q_per_col = (num_mc_questions + 1) // 2
    col = q_idx // q_per_col
    row = q_idx % q_per_col
    qx  = MARGIN_L + col * COL_W
    qy  = SEC1_FIRST_Q_Y - row * BUBBLE_SP_Y
    bx  = qx + BUBBLE_OFFSET + choice_idx * BUBBLE_SP_X
    return bx, qy

def sec1_divider_y(num_mc_questions):
    q_per_col = (num_mc_questions + 1) // 2
    return SEC1_FIRST_Q_Y - (q_per_col - 1) * BUBBLE_SP_Y - 24

# ── Section 2 ─────────────────────────────────────────────────────────────────
SK_Q = 10

def sec2_title_y(num_mc_questions):
    return sec1_divider_y(num_mc_questions) - 18

def sec2_bubble(q_idx, choice_idx, num_mc_questions):
    sk_per_col = (SK_Q + 1) // 2
    col = q_idx // sk_per_col
    row = q_idx % sk_per_col
    s2y = sec2_title_y(num_mc_questions)
    qx  = MARGIN_L + col * COL_W
    qy  = s2y - 42 - row * BUBBLE_SP_Y
    bx  = qx + BUBBLE_OFFSET + choice_idx * BUBBLE_SP_X
    return bx, qy
