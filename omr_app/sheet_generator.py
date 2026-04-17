"""
OMR Sheet Generator — all coordinates from layout.py.
"""
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

from layout import (
    PAGE_W, PAGE_H, MARGIN_L, BUBBLE_R, BUBBLE_SP_X, BUBBLE_SP_Y,
    COL_W, BUBBLE_OFFSET, Q_NUM_OFFSET,
    LOGO_W, LOGO_H, LOGO_X, LOGO_Y, TITLE_X, TITLE_Y,
    FIELDS_Y1, BUBBLE_ROW_Y, SEPARATOR_Y,
    GRADO_LABEL_X, GRADO_B1_X, GRADO_VALUES, grado_bubble_x,
    GRUPO_LABEL_X, GRUPO_B1_X, GRUPO_VALUES, grupo_bubble_x,
    FOLIO_LABEL_X, FOLIO_B1_X, FOLIO_VALUES, FOLIO_ROW1_Y, FOLIO_ROW2_Y, folio_bubble_x,
    SEC1_TITLE_Y, SEC1_RANGE_Y, SEC1_HDR_Y, SEC1_FIRST_Q_Y,
    sec1_divider_y, SK_Q, sec2_title_y,
    MARKER_SIZE, MARKERS,
)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")
FONT_SIZE  = 12


def _bubble(c, x, y):
    c.setLineWidth(1.0)
    c.setStrokeColor(colors.black)
    c.setFillColor(colors.white)
    c.circle(x, y, BUBBLE_R, stroke=1, fill=1)
    c.setFillColor(colors.black)


def _draw_page(c, num_mc_questions, exam_name):

    # ── Corner markers ────────────────────────────────────────────────────────
    for mx, my in MARKERS:
        c.setFillColor(colors.black)
        c.rect(mx - MARKER_SIZE/2, my - MARKER_SIZE/2,
               MARKER_SIZE, MARKER_SIZE, stroke=0, fill=1)

    # ── Logo ──────────────────────────────────────────────────────────────────
    if os.path.exists(LOGO_PATH):
        try:
            c.drawImage(ImageReader(LOGO_PATH), LOGO_X, LOGO_Y,
                        width=LOGO_W, height=LOGO_H, mask='auto')
        except Exception:
            pass

    # ── Exam title — centered ─────────────────────────────────────────────────
    c.setFillColor(colors.HexColor("#1B2A6B"))
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(TITLE_X, TITLE_Y, exam_name)

    # ── Field row 1: Nombre, Fecha, Escuela ───────────────────────────────────
    c.setFillColor(colors.black)
    c.setFont("Helvetica", FONT_SIZE)
    for fx, text in [
        (MARGIN_L,           "Nombre: ___________________________"),
        (MARGIN_L + 3.3*72,  "Fecha: ____________"),
        (MARGIN_L + 5.0*72,  "Escuela: ___________________"),
    ]:
        c.drawString(fx, FIELDS_Y1, text)

    # ── Bubble zone: Grado | Grupo | Folio ───────────────────────────────────
    HDR_Y = BUBBLE_ROW_Y + BUBBLE_R + 7   # label above Grado/Grupo bubbles
    MID_Y = BUBBLE_ROW_Y - 4              # vertical center for section labels

    # Grado
    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.setFillColor(colors.black)
    c.drawString(GRADO_LABEL_X, MID_Y, "Grado")
    for i, val in enumerate(GRADO_VALUES):
        bx = grado_bubble_x(i)
        _bubble(c, bx, BUBBLE_ROW_Y)
        c.setFont("Helvetica", FONT_SIZE)
        c.setFillColor(colors.black)
        c.drawCentredString(bx, HDR_Y, val)

    # Grupo
    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.setFillColor(colors.black)
    c.drawString(GRUPO_LABEL_X, MID_Y, "Grupo")
    for i, val in enumerate(GRUPO_VALUES):
        bx = grupo_bubble_x(i)
        _bubble(c, bx, BUBBLE_ROW_Y)
        c.setFont("Helvetica", FONT_SIZE)
        c.setFillColor(colors.black)
        c.drawCentredString(bx, HDR_Y, val)

    # Folio — two rows of 9 bubbles (1–9)
    FOLIO_HDR_Y  = FOLIO_ROW1_Y + BUBBLE_R + 7   # digit labels above row 1
    folio_n_edge = FOLIO_B1_X - BUBBLE_R - 4      # right edge for N1/N2 labels

    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.setFillColor(colors.black)
    c.drawString(FOLIO_LABEL_X, MID_Y, "Folio")

    # Digit labels 1–9 (above row 1 only, shared reference for both rows)
    for i, val in enumerate(FOLIO_VALUES):
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.black)
        c.drawCentredString(folio_bubble_x(i), FOLIO_HDR_Y, val)

    # Row 1 — Número 1
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.black)
    c.drawRightString(folio_n_edge, FOLIO_ROW1_Y - 4, "N1")
    for i in range(len(FOLIO_VALUES)):
        _bubble(c, folio_bubble_x(i), FOLIO_ROW1_Y)

    # Row 2 — Número 2
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.black)
    c.drawRightString(folio_n_edge, FOLIO_ROW2_Y - 4, "N2")
    for i in range(len(FOLIO_VALUES)):
        _bubble(c, folio_bubble_x(i), FOLIO_ROW2_Y)

    # ── Separator ─────────────────────────────────────────────────────────────
    c.setStrokeColor(colors.HexColor("#CCCCCC"))
    c.setLineWidth(0.4)
    c.line(MARGIN_L, SEPARATOR_Y, PAGE_W - MARGIN_L, SEPARATOR_Y)

    # ── Section 1 ─────────────────────────────────────────────────────────────
    q_per_col = (num_mc_questions + 1) // 2

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.drawString(MARGIN_L, SEC1_TITLE_Y,
                 f"Sección 1 — Opción Múltiple  ({num_mc_questions} preguntas, A–D)")

    for col_idx in range(2):
        q_start = col_idx * q_per_col + 1
        q_end   = min((col_idx+1) * q_per_col, num_mc_questions)
        if q_start > num_mc_questions:
            break
        lx = MARGIN_L + col_idx * COL_W
        c.setFillColor(colors.grey)
        c.setFont("Helvetica-Oblique", FONT_SIZE - 2)
        c.drawString(lx, SEC1_RANGE_Y, f"({q_start}–{q_end})")
        for j, ch in enumerate(["A","B","C","D"]):
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", FONT_SIZE)
            c.drawCentredString(lx + BUBBLE_OFFSET + j*BUBBLE_SP_X, SEC1_HDR_Y, ch)

    for i in range(num_mc_questions):
        col = i // q_per_col
        row = i % q_per_col
        qx  = MARGIN_L + col * COL_W
        qy  = SEC1_FIRST_Q_Y - row * BUBBLE_SP_Y
        c.setFillColor(colors.black)
        c.setFont("Helvetica", FONT_SIZE)
        c.drawRightString(qx + Q_NUM_OFFSET, qy - 4, f"{i+1}.")
        for j in range(4):
            _bubble(c, qx + BUBBLE_OFFSET + j*BUBBLE_SP_X, qy)

    # ── Divider ───────────────────────────────────────────────────────────────
    div_y = sec1_divider_y(num_mc_questions)
    c.setStrokeColor(colors.grey)
    c.setLineWidth(0.5)
    c.line(MARGIN_L - 5, div_y, PAGE_W - MARGIN_L + 5, div_y)

    # ── Section 2 ─────────────────────────────────────────────────────────────
    s2y        = sec2_title_y(num_mc_questions)
    sk_per_col = (SK_Q + 1) // 2

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.drawString(MARGIN_L, s2y,
                 "Sección 2 — Autoconocimiento  (10 preguntas, escala 1–5)")

    for col_idx in range(2):
        q_start = col_idx * sk_per_col + 1
        q_end   = min((col_idx+1) * sk_per_col, SK_Q)
        if q_start > SK_Q:
            break
        lx = MARGIN_L + col_idx * COL_W
        c.setFillColor(colors.grey)
        c.setFont("Helvetica-Oblique", FONT_SIZE - 2)
        c.drawString(lx, s2y - 14, f"({q_start}–{q_end})")
        for j, lbl in enumerate(["1","2","3","4","5"]):
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", FONT_SIZE)
            c.drawCentredString(lx + BUBBLE_OFFSET + j*BUBBLE_SP_X, s2y - 28, lbl)

    for i in range(SK_Q):
        col = i // sk_per_col
        row = i % sk_per_col
        qx  = MARGIN_L + col * COL_W
        qy  = s2y - 42 - row * BUBBLE_SP_Y
        c.setFillColor(colors.black)
        c.setFont("Helvetica", FONT_SIZE)
        c.drawRightString(qx + Q_NUM_OFFSET, qy - 4, f"{i+1}.")
        for j in range(5):
            _bubble(c, qx + BUBBLE_OFFSET + j*BUBBLE_SP_X, qy)

    # ── Footer ────────────────────────────────────────────────────────────────
    c.setFillColor(colors.grey)
    c.setFont("Helvetica", 9)
    c.drawCentredString(PAGE_W/2, 20,
        "Use lápiz o bolígrafo negro. Rellene completamente el círculo.")


def generate_omr_sheet(output_path, num_mc_questions=20, exam_name="Examen"):
    c = canvas.Canvas(output_path, pagesize=letter)
    _draw_page(c, num_mc_questions, exam_name)
    c.save()
    return output_path


def generate_batch_sheets(output_path, num_mc_questions=20,
                          num_students=30, exam_name="Examen"):
    c = canvas.Canvas(output_path, pagesize=letter)
    for _ in range(num_students):
        _draw_page(c, num_mc_questions, exam_name)
        c.showPage()
    c.save()
    return output_path
