"""
OMR Scanner — all grid coordinates from layout.py.
"""
import cv2
import numpy as np
from pdf2image import convert_from_path
from dataclasses import dataclass
from typing import Optional

from layout import (
    PAGE_W, PAGE_H, BUBBLE_R,
    GRADO_VALUES, GRUPO_VALUES, BUBBLE_ROW_Y,
    grado_bubble_x, grupo_bubble_x,
    FOLIO_VALUES, FOLIO_ROW1_Y, FOLIO_ROW2_Y, folio_bubble_x,
    SK_Q, sec1_bubble, sec2_bubble,
    SEC3_WORDS, sec3_bubble,
    MARKER_OFFSET, MARKER_SIZE,
)

WARP_W     = 850
WARP_H     = 1100
SCALE_X    = WARP_W / PAGE_W
SCALE_Y    = WARP_H / PAGE_H
SCAN_R     = 9
FILL_THRESHOLD = 0.28   # works for pencil (~0.40) and pen (~0.60+); noise/smudge ≈ 0.12


@dataclass
class ScanResult:
    page_num:        int
    folio:           str             # scanned 2-digit folio e.g. '37', '?5', '??'
    grado:           Optional[str]   # '1','2','3', '?' or None
    grupo:           Optional[str]   # 'A'–'F', '?' or None
    mc_answers:      list
    sk_answers:      list
    word_selections: list            # list[bool], one per SEC3_WORDS entry
    confidence:      float
    error:           Optional[str] = None


def _pt_to_px(x_pt, y_pt):
    return int(x_pt * SCALE_X), int((PAGE_H - y_pt) * SCALE_Y)


def _bub(x_pt, y_pt):
    px, py = _pt_to_px(x_pt, y_pt)
    return (px, py, SCAN_R)


def pdf_to_images(pdf_path, dpi=200):
    """Load all pages at once — only safe for small PDFs (≤ ~30 pages)."""
    return [np.array(img.convert("RGB"))
            for img in convert_from_path(pdf_path, dpi=dpi)]


def _pdf_page_count(pdf_path):
    from pdf2image import pdfinfo_from_path
    return pdfinfo_from_path(pdf_path)["Pages"]


SCAN_BATCH = 50   # pages rendered at a time — keeps peak RAM ≈ 550 MB


def _preprocess(img_rgb):
    gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blur, 0, 255,
                              cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    return gray, thresh


def _find_corners(thresh, shape):
    """
    For each corner quadrant pick the largest-area square-ish contour.
    The real corner markers are always the biggest dark squares in their quadrant.

    Improvements over the naive bounding-box approach:
    - Use contour moments for sub-pixel-accurate centroid (avoids the ~½-bubble
      drift caused by asymmetric contour borders on a bounding-box midpoint).
    - Tighter aspect-ratio filter (0.60–1.70) rejects horizontal lines / text.
    - Coarse area filter tuned to the printed marker size at 150–300 DPI.
    - Corner zone widened to 20 % so slightly-skewed scans still find all four.
    """
    h, w = shape[:2]
    # Expected marker area at 200 Dpi: (18pt * 200/72)^2 ≈ 2 500 px²
    # Allow ×0.15 – ×4 range for variation in scanner exposure / DPI
    MIN_AREA = 300
    MAX_AREA = 12_000

    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL,
                                   cv2.CHAIN_APPROX_SIMPLE)
    quad = {'tl': [], 'tr': [], 'bl': [], 'br': []}
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if not (MIN_AREA <= area <= MAX_AREA):
            continue
        bx, by, cw, ch = cv2.boundingRect(cnt)
        aspect = cw / max(ch, 1)
        if not (0.60 < aspect < 1.70):          # must be roughly square
            continue

        # Moment-based centroid is more accurate than bounding-box midpoint
        M = cv2.moments(cnt)
        if M["m00"] == 0:
            continue
        cx = int(M["m10"] / M["m00"])
        cy = int(M["m01"] / M["m00"])

        in_l = cx < w * 0.20
        in_r = cx > w * 0.80
        in_t = cy < h * 0.20
        in_b = cy > h * 0.80
        if in_l and in_t:
            quad['tl'].append((cx, cy, area))
        if in_r and in_t:
            quad['tr'].append((cx, cy, area))
        if in_l and in_b:
            quad['bl'].append((cx, cy, area))
        if in_r and in_b:
            quad['br'].append((cx, cy, area))

    corners = []
    for key in ('tl', 'tr', 'bl', 'br'):
        if not quad[key]:
            return None
        best = max(quad[key], key=lambda c: c[2])
        corners.append((best[0], best[1]))
    return corners  # [tl, tr, bl, br]


def _warp(img, corners):
    tl, tr, bl, br = corners
    src = np.float32([tl, tr, bl, br])
    # Map each marker center to its expected pixel position in the full-page
    # coordinate system so that _pt_to_px() works correctly after the warp.
    mx = int(MARKER_OFFSET * SCALE_X)
    my = int(MARKER_OFFSET * SCALE_Y)
    dst = np.float32([
        [mx,          my],
        [WARP_W - mx, my],
        [mx,          WARP_H - my],
        [WARP_W - mx, WARP_H - my],
    ])
    return cv2.warpPerspective(img,
                               cv2.getPerspectiveTransform(src, dst),
                               (WARP_W, WARP_H))


def _fill_ratio(gray, x, y, r):
    mask = np.zeros(gray.shape, dtype=np.uint8)
    cv2.circle(mask, (x, y), max(r-1, 3), 255, -1)
    total = cv2.countNonZero(mask)
    if total == 0:
        return 0.0
    _, inv = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV)
    return cv2.countNonZero(cv2.bitwise_and(inv, inv, mask=mask)) / total


def _best_in_row(gray, bubbles, threshold=FILL_THRESHOLD):
    """Pick most-filled bubble. Returns index or None if below threshold."""
    ratios = [_fill_ratio(gray, x, y, r) for x, y, r in bubbles]
    best   = max(ratios)
    return ratios.index(best) if best >= threshold else None


def _best_with_ambiguity(gray, bubbles, values, threshold=FILL_THRESHOLD):
    """
    Like _best_in_row but also detects if multiple bubbles are filled.
    Returns the value string, '?' if ambiguous, or None if blank.
    """
    ratios   = [_fill_ratio(gray, x, y, r) for x, y, r in bubbles]
    filled   = [i for i, r in enumerate(ratios) if r >= threshold]
    if len(filled) == 0:
        return None
    if len(filled) > 1:
        return '?'
    return values[filled[0]]


def _mc_grid(n):
    return [[_bub(*sec1_bubble(i, j, n)) for j in range(4)] for i in range(n)]


def _sk_grid(n):
    return [[_bub(*sec2_bubble(i, j, n)) for j in range(5)] for i in range(SK_Q)]


def _grado_row():
    return [_bub(grado_bubble_x(i), BUBBLE_ROW_Y) for i in range(len(GRADO_VALUES))]


def _grupo_row():
    return [_bub(grupo_bubble_x(i), BUBBLE_ROW_Y) for i in range(len(GRUPO_VALUES))]


def _folio_row(row_y):
    return [_bub(folio_bubble_x(i), row_y) for i in range(len(FOLIO_VALUES))]


def _word_row(n):
    """One bubble per word — checked independently (multi-select)."""
    return [_bub(*sec3_bubble(i, n)) for i in range(len(SEC3_WORDS))]


MC_CHOICES = ['A','B','C','D']
SK_CHOICES  = [1, 2, 3, 4, 5]


def scan_page(img_rgb, num_mc_questions):
    gray, thresh = _preprocess(img_rgb)
    corners = _find_corners(thresh, img_rgb.shape)
    if corners is not None:
        wgray = cv2.cvtColor(_warp(img_rgb, corners), cv2.COLOR_RGB2GRAY)
    else:
        wgray = cv2.resize(gray, (WARP_W, WARP_H))

    # MC answers
    mc_answers, filled = [], 0
    for row in _mc_grid(num_mc_questions):
        idx = _best_in_row(wgray, row)
        mc_answers.append(MC_CHOICES[idx] if idx is not None else None)
        if idx is not None:
            filled += 1

    # SK answers
    sk_answers = []
    for row in _sk_grid(num_mc_questions):
        idx = _best_in_row(wgray, row)
        sk_answers.append(SK_CHOICES[idx] if idx is not None else None)

    # Grado, Grupo — ambiguity-aware
    grado = _best_with_ambiguity(wgray, _grado_row(), GRADO_VALUES)
    grupo = _best_with_ambiguity(wgray, _grupo_row(), GRUPO_VALUES)

    # Folio — two rows, each picks one digit 0–9
    d1 = _best_with_ambiguity(wgray, _folio_row(FOLIO_ROW1_Y), FOLIO_VALUES)
    d2 = _best_with_ambiguity(wgray, _folio_row(FOLIO_ROW2_Y), FOLIO_VALUES)
    folio_str = f"{d1 or '?'}{d2 or '?'}"

    # Word selections — each bubble is independent (multi-select allowed)
    word_selections = [
        _fill_ratio(wgray, x, y, r) >= FILL_THRESHOLD
        for x, y, r in _word_row(num_mc_questions)
    ]

    return ScanResult(
        page_num        = 0,
        folio           = folio_str,
        grado           = grado,
        grupo           = grupo,
        mc_answers      = mc_answers,
        sk_answers      = sk_answers,
        word_selections = word_selections,
        confidence      = filled / max(num_mc_questions, 1),
    )


def scan_single_page_from_pdf(pdf_path, page_num, num_mc_questions):
    """Re-scan a single page (1-indexed) from a multi-page PDF."""
    images = convert_from_path(pdf_path, dpi=200,
                               first_page=page_num, last_page=page_num)
    if not images:
        raise ValueError(f"No se pudo leer la página {page_num} del PDF.")
    img_rgb        = np.array(images[0].convert("RGB"))
    result         = scan_page(img_rgb, num_mc_questions)
    result.page_num = page_num
    return result


def scan_pdf(pdf_path, num_mc_questions, progress_callback=None):
    """
    Scan all pages using batched rendering so peak RAM stays bounded
    (≈ SCAN_BATCH × 11 MB) regardless of total page count.
    """
    total   = _pdf_page_count(pdf_path)
    results = []

    for batch_start in range(1, total + 1, SCAN_BATCH):
        batch_end = min(batch_start + SCAN_BATCH - 1, total)
        images    = convert_from_path(pdf_path, dpi=200,
                                      first_page=batch_start,
                                      last_page=batch_end)
        for j, pil_img in enumerate(images):
            page_num = batch_start + j
            if progress_callback:
                progress_callback(page_num - 1, total)
            img = np.array(pil_img.convert("RGB"))
            try:
                result          = scan_page(img, num_mc_questions)
                result.page_num = page_num
            except Exception as e:
                result = ScanResult(
                    page_num        = page_num,
                    folio           = "??",
                    grado           = None,
                    grupo           = None,
                    mc_answers      = [None] * num_mc_questions,
                    sk_answers      = [None] * SK_Q,
                    word_selections = [False] * len(SEC3_WORDS),
                    confidence      = 0.0,
                    error           = str(e),
                )
            results.append(result)
    return results
