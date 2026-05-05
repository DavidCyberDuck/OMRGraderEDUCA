"""
Excel Exporter — results to .xlsx with 4 sheets and charts.
Column order: Folio → Grado → Grupo → Score...
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
import datetime

C_HDR_BG = "2C3E50"
C_HDR_FG = "FFFFFF"
C_ACCENT = "3498DB"
C_GREEN  = "D5F5E3"
C_RED    = "FADBD8"
C_ALT    = "F8F9FA"
C_BORDER = "BDC3C7"


def _hdr(cell, bg=C_HDR_BG, fg=C_HDR_FG):
    cell.font      = Font(bold=True, color=fg, name="Arial", size=10)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
    cell.border    = _border()


def _border():
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)


def _cw(ws, col, w):
    ws.column_dimensions[get_column_letter(col)].width = w


def export_to_excel(grade_results, answer_key, exam_name, output_path,
                    student_db=None):
    wb = Workbook()
    _summary(wb, grade_results, exam_name, student_db)
    _detail(wb, grade_results, answer_key)
    _sk_sheet(wb, grade_results, student_db)
    _words_sheet(wb, grade_results)
    _charts(wb, grade_results, student_db)
    _clave_sheet(wb, answer_key, exam_name)
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    wb.save(output_path)
    return output_path


def _summary(wb, results, exam_name, student_db=None):
    ws = wb.create_sheet("Resumen")
    ws.sheet_view.showGridLines = False

    has_db = student_db is not None
    n_cols = 10 if has_db else 9
    last_col_ltr = get_column_letter(n_cols)

    ws.merge_cells(f"A1:{last_col_ltr}1")
    ws["A1"].value = f"Resultados — {exam_name}"
    ws["A1"].font  = Font(bold=True, size=14, color=C_HDR_FG, name="Arial")
    ws["A1"].fill  = PatternFill("solid", fgColor=C_HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col_ltr}2")
    ws["A2"].value = f"Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font  = Font(italic=True, size=9, color="888888", name="Arial")
    ws["A2"].alignment = Alignment(horizontal="right")

    if has_db:
        headers = ["Folio","Nombre","Grado","Grupo","Puntaje","Total",
                   "Porcentaje (%)","Prom. Autoconoc.","Confianza (%)","Estado"]
        pct_col = 7
        col_widths = [8,22,8,8,10,8,16,18,14,10]
    else:
        headers = ["Folio","Grado","Grupo","Puntaje","Total",
                   "Porcentaje (%)","Prom. Autoconoc.","Confianza (%)","Estado"]
        pct_col = 6
        col_widths = [8,8,8,10,8,16,18,14,10]

    for col, h in enumerate(headers, 1):
        _hdr(ws.cell(row=4, column=col, value=h))
    ws.row_dimensions[4].height = 32

    C_YELLOW = "FFF3CD"

    for r, gr in enumerate(results, 5):
        bg = C_ALT if r % 2 == 0 else "FFFFFF"
        if has_db:
            nombre = student_db.get(str(gr.folio), "")
            row_data = [
                gr.folio, nombre, gr.grado or "?", gr.grupo or "?",
                gr.score, gr.total, gr.percentage,
                gr.sk_average if gr.sk_average is not None else "N/A",
                round(gr.confidence * 100, 0),
                "⚠ Error" if gr.error else "OK",
            ]
        else:
            row_data = [
                gr.folio, gr.grado or "?", gr.grupo or "?",
                gr.score, gr.total, gr.percentage,
                gr.sk_average if gr.sk_average is not None else "N/A",
                round(gr.confidence * 100, 0),
                "⚠ Error" if gr.error else "OK",
            ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _border()

        # Yellow Nombre cell when folio not found in db
        if has_db and not student_db.get(str(gr.folio)):
            ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=C_YELLOW)

        # Color-code percentage column
        pct_cell = ws.cell(row=r, column=pct_col)
        pct_cell.fill = PatternFill("solid",
                        fgColor=C_GREEN if gr.percentage >= 70 else C_RED)

    last = 4 + len(results)
    pct_letter = get_column_letter(pct_col)
    sr   = last + 2
    ws.cell(row=sr, column=1, value="Estadísticas").font = Font(bold=True, name="Arial")
    for i, (lbl, fml) in enumerate([
        ("Promedio",         f"=AVERAGE({pct_letter}5:{pct_letter}{last})"),
        ("Máximo",           f"=MAX({pct_letter}5:{pct_letter}{last})"),
        ("Mínimo",           f"=MIN({pct_letter}5:{pct_letter}{last})"),
        ("Aprobados (≥70%)", f'=COUNTIF({pct_letter}5:{pct_letter}{last},">=70")'),
    ]):
        ws.cell(row=sr+1+i, column=1, value=lbl).font = Font(bold=True, name="Arial", size=10)
        ws.cell(row=sr+1+i, column=2, value=fml).font = Font(name="Arial", size=10)

    for col, w in enumerate(col_widths, 1):
        _cw(ws, col, w)

    ws.conditional_formatting.add(f"{pct_letter}5:{pct_letter}{last}", ColorScaleRule(
        start_type="num", start_value=0,   start_color="E74C3C",
        mid_type="num",   mid_value=70,    mid_color="F39C12",
        end_type="num",   end_value=100,   end_color="2ECC71",
    ))


def _detail(wb, results, answer_key):
    ws  = wb.create_sheet("Detalle Preguntas")
    ws.sheet_view.showGridLines = False
    n_q = len(answer_key)

    ws.merge_cells(f"A1:{get_column_letter(n_q+5)}1")
    ws["A1"].value = "Detalle de Respuestas — Sección 1"
    ws["A1"].font  = Font(bold=True, size=13, color=C_HDR_FG, name="Arial")
    ws["A1"].fill  = PatternFill("solid", fgColor=C_HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for col, val in enumerate(["Clave","Folio","Grado","Grupo"] +
                               answer_key, 1):
        cell = ws.cell(row=2, column=col, value=val)
        cell.font = Font(bold=True,
                         color=C_HDR_FG if col > 4 else "000000",
                         name="Arial", size=9)
        if col > 4:
            cell.fill = PatternFill("solid", fgColor=C_ACCENT)
        cell.alignment = Alignment(horizontal="center")

    for col, val in enumerate(["","Folio","Grado","Grupo"] +
                               [f"P{q}" for q in range(1, n_q+1)], 1):
        cell = ws.cell(row=3, column=col, value=val)
        cell.font = Font(bold=True, name="Arial", size=8)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill("solid", fgColor="ECF0F1")

    for r, gr in enumerate(results, 4):
        ws.cell(row=r, column=1, value=gr.page_num)
        ws.cell(row=r, column=2, value=gr.folio)
        ws.cell(row=r, column=3, value=gr.grado or "?")
        ws.cell(row=r, column=4, value=gr.grupo or "?")
        for q, (ans, ok) in enumerate(zip(gr.mc_answers, gr.mc_correct)):
            cell = ws.cell(row=r, column=q+5, value=ans or "-")
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill("solid", fgColor=C_GREEN if ok else C_RED)

    last = 3 + len(results)
    acc  = last + 2
    ws.cell(row=acc, column=1, value="Aciertos %").font = Font(bold=True, name="Arial", size=9)
    for q in range(n_q):
        correct = sum(1 for gr in results
                      if q < len(gr.mc_correct) and gr.mc_correct[q])
        pct  = round(correct / max(len(results), 1) * 100, 1)
        cell = ws.cell(row=acc, column=q+5, value=pct)
        cell.font = Font(name="Arial", size=9)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill("solid",
                    fgColor=C_GREEN if pct >= 80 else (C_RED if pct < 50 else "FFFFFF"))

    for col in range(1, n_q+6):
        _cw(ws, col, 5.5)
    for col, w in enumerate([6,8,7,7], 1):
        _cw(ws, col, w)


def _sk_sheet(wb, results, student_db=None):
    ws = wb.create_sheet("Autoconocimiento")
    ws.sheet_view.showGridLines = False

    has_db  = student_db is not None
    sk_off  = 4 if has_db else 4   # SK1 starts at col 4 (same either way: Folio,Nombre?,Grado,Grupo)
    avg_col = 15 if has_db else 14
    n_cols  = avg_col
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    ws["A1"].value = "Sección 2 — Autoconocimiento"
    ws["A1"].font  = Font(bold=True, size=13, color=C_HDR_FG, name="Arial")
    ws["A1"].fill  = PatternFill("solid", fgColor=C_HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    if has_db:
        headers = ["Folio","Nombre","Grado","Grupo"] + \
                  [f"SK{i}" for i in range(1,11)] + ["Promedio"]
        sk_off  = 5
        avg_col = 15
        meta_widths = [8,22,8,8]
    else:
        headers = ["Folio","Grado","Grupo"] + \
                  [f"SK{i}" for i in range(1,11)] + ["Promedio"]
        sk_off  = 4
        avg_col = 14
        meta_widths = [8,8,8]

    for col, h in enumerate(headers, 1):
        _hdr(ws.cell(row=2, column=col, value=h), bg=C_ACCENT)

    C_YELLOW = "FFF3CD"
    for r, gr in enumerate(results, 3):
        bg = C_ALT if r % 2 == 0 else "FFFFFF"
        if has_db:
            nombre = student_db.get(str(gr.folio), "")
            meta   = [gr.folio, nombre, gr.grado or "?", gr.grupo or "?"]
        else:
            meta   = [gr.folio, gr.grado or "?", gr.grupo or "?"]
        for col, val in enumerate(meta, 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="center")
        if has_db and not student_db.get(str(gr.folio)):
            ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=C_YELLOW)

        for i, val in enumerate(gr.sk_answers):
            cell = ws.cell(row=r, column=sk_off + i,
                           value=val if val is not None else "-")
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(name="Arial", size=10)
            cell.fill = PatternFill("solid", fgColor=bg)
        avg = ws.cell(row=r, column=avg_col,
                      value=gr.sk_average if gr.sk_average else "-")
        avg.alignment = Alignment(horizontal="center")
        avg.font = Font(name="Arial", size=10, bold=True)

    for col, w in enumerate(meta_widths + [7]*10 + [10], 1):
        _cw(ws, col, w)


def _charts(wb, results, student_db=None):
    ws = wb.create_sheet("Gráficas")
    ws.sheet_view.showGridLines = False

    # ── Title banner ──────────────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    ws["A1"].value     = "Gráficas de Resultados"
    ws["A1"].font      = Font(bold=True, size=14, color=C_HDR_FG, name="Arial")
    ws["A1"].fill      = PatternFill("solid", fgColor=C_HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:G2")
    ws["A2"].value     = f"Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font      = Font(italic=True, size=9, color="888888", name="Arial")
    ws["A2"].alignment = Alignment(horizontal="right")

    # ── Individual data table ─────────────────────────────────────────────────
    has_db = student_db is not None
    IND_HDR = 4
    ind_headers = (["Folio", "Nombre", "Grado", "Grupo", "Puntaje MC (%)", "Prom. Autoconoc."]
                   if has_db else
                   ["Folio", "Grado", "Grupo", "Puntaje MC (%)", "Prom. Autoconoc."])
    for col, h in enumerate(ind_headers, 1):
        _hdr(ws.cell(row=IND_HDR, column=col, value=h), bg=C_ACCENT)
    ws.row_dimensions[IND_HDR].height = 26

    for i, gr in enumerate(results):
        r      = IND_HDR + 1 + i
        bg     = C_ALT if r % 2 == 0 else "FFFFFF"
        nombre = (student_db or {}).get(str(gr.folio), "")
        row_data = ([gr.folio, nombre, gr.grado or "?", gr.grupo or "?",
                     gr.percentage, gr.sk_average or 0] if has_db else
                    [gr.folio, gr.grado or "?", gr.grupo or "?",
                     gr.percentage, gr.sk_average or 0])
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _border()

    if has_db:
        for col, w in enumerate([8, 22, 8, 8, 16, 16], 1):
            _cw(ws, col, w)
    else:
        for col, w in enumerate([8, 8, 8, 16, 16], 1):
            _cw(ws, col, w)

    # ── Group summary table ───────────────────────────────────────────────────
    # Aggregate by (Grado, Grupo), sorted so charts read in natural order
    from collections import defaultdict
    groups = defaultdict(lambda: {"pct": [], "sk": []})
    for gr in results:
        key = (gr.grado or "?", gr.grupo or "?")
        groups[key]["pct"].append(gr.percentage)
        if gr.sk_average is not None:
            groups[key]["sk"].append(gr.sk_average)

    sorted_groups = sorted(groups.items(), key=lambda x: (x[0][0], x[0][1]))

    GRP_HDR = IND_HDR + len(results) + 3   # two blank rows gap
    grp_label_col = 1   # "Grado X / Grupo Y"
    grp_n_col     = 2
    grp_pct_col   = 3
    grp_sk_col    = 4

    # Section label
    lbl = ws.cell(row=GRP_HDR - 1, column=1,
                  value="Resumen por Grado y Grupo")
    lbl.font = Font(bold=True, size=11, color=C_HDR_FG, name="Arial")
    lbl.fill = PatternFill("solid", fgColor=C_HDR_BG)
    ws.merge_cells(start_row=GRP_HDR - 1, start_column=1,
                   end_row=GRP_HDR - 1, end_column=4)
    ws.cell(row=GRP_HDR - 1, column=1).alignment = Alignment(
        horizontal="center", vertical="center")
    ws.row_dimensions[GRP_HDR - 1].height = 22

    for col, h in enumerate(["Grado / Grupo", "N Alumnos",
                              "Prom. MC (%)", "Prom. Autoconoc."], 1):
        _hdr(ws.cell(row=GRP_HDR, column=col, value=h), bg=C_ACCENT)
    ws.row_dimensions[GRP_HDR].height = 26

    GRP_FIRST = GRP_HDR + 1
    for i, ((grado, grupo), vals) in enumerate(sorted_groups):
        r   = GRP_FIRST + i
        bg  = C_ALT if r % 2 == 0 else "FFFFFF"
        avg_pct = round(sum(vals["pct"]) / len(vals["pct"]), 1) if vals["pct"] else 0
        avg_sk  = round(sum(vals["sk"])  / len(vals["sk"]),  2) if vals["sk"]  else "-"
        label   = f"Grado {grado} / Grupo {grupo}"
        for col, val in enumerate([label, len(vals["pct"]), avg_pct, avg_sk], 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _border()
        ws.cell(row=r, column=1).alignment = Alignment(
            horizontal="left", vertical="center")

    GRP_LAST = GRP_FIRST + len(sorted_groups) - 1

    for col, w in enumerate([22, 12, 16, 16], 1):
        _cw(ws, col, w)

    # ── Chart 1: Prom. MC (%) por grupo ──────────────────────────────────────
    bar = BarChart()
    bar.type   = "col"
    bar.title  = "Sección 1 — Puntaje Promedio por Grupo (%)"
    bar.style  = 2
    bar.y_axis.title = "Porcentaje (%)"
    bar.x_axis.title = "Grupo"
    bar.y_axis.scaling.min = 0
    bar.y_axis.scaling.max = 100
    bar.height = 14
    bar.width  = 22
    bar.add_data(Reference(ws, min_col=grp_pct_col,
                           min_row=GRP_HDR, max_row=GRP_LAST),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=grp_label_col,
                                 min_row=GRP_FIRST, max_row=GRP_LAST))
    bar.series[0].graphicalProperties.solidFill = C_ACCENT
    bar.series[0].graphicalProperties.line.solidFill = C_ACCENT
    ws.add_chart(bar, "F4")

    # ── Chart 2: Prom. Autoconoc. por grupo ──────────────────────────────────
    bar2 = BarChart()
    bar2.type   = "col"
    bar2.title  = "Sección 2 — Promedio de Autoconocimiento por Grupo"
    bar2.style  = 2
    bar2.y_axis.title = "Promedio (1–5)"
    bar2.x_axis.title = "Grupo"
    bar2.y_axis.scaling.min = 0
    bar2.y_axis.scaling.max = 5
    bar2.height = 14
    bar2.width  = 22
    bar2.add_data(Reference(ws, min_col=grp_sk_col,
                            min_row=GRP_HDR, max_row=GRP_LAST),
                  titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=grp_label_col,
                                  min_row=GRP_FIRST, max_row=GRP_LAST))
    bar2.series[0].graphicalProperties.solidFill = "E74C3C"
    bar2.series[0].graphicalProperties.line.solidFill = "E74C3C"
    ws.add_chart(bar2, "F27")


def _words_sheet(wb, results):
    from layout import SEC3_WORDS
    ws = wb.create_sheet("Palabras ABE")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:B1")
    ws["A1"].value = "Sección 3 — Palabras ABE"
    ws["A1"].font  = Font(bold=True, size=13, color=C_HDR_FG, name="Arial")
    ws["A1"].fill  = PatternFill("solid", fgColor=C_HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(["Palabra", "Selecciones"], 1):
        _hdr(ws.cell(row=3, column=col, value=h))
    ws.row_dimensions[3].height = 28

    for idx, word in enumerate(SEC3_WORDS):
        count = sum(
            1 for gr in results
            if gr.word_selections and idx < len(gr.word_selections)
               and gr.word_selections[idx]
        )
        row = idx + 4
        bg  = C_ALT if row % 2 == 0 else "FFFFFF"
        for col, val in enumerate([word, count], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(
                horizontal="left" if col == 1 else "center",
                vertical="center")
            cell.border = _border()

    _cw(ws, 1, 22)
    _cw(ws, 2, 14)

    bar = BarChart()
    bar.type   = "bar"   # horizontal bars — easier to read with long word labels
    bar.title  = "Palabras ABE — Frecuencia"
    bar.style  = 10
    bar.x_axis.title = "Selecciones"
    bar.height = 14
    bar.width  = 18
    last_word_row = 3 + len(SEC3_WORDS)
    bar.add_data(Reference(ws, min_col=2, min_row=3, max_row=last_word_row),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=4, max_row=last_word_row))
    ws.add_chart(bar, "D3")


def _clave_sheet(wb, answer_key, exam_name):
    ws = wb.create_sheet("Clave de Respuestas")
    ws.sheet_view.showGridLines = False
    n        = len(answer_key)
    last_col = get_column_letter(n + 1)

    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"].value     = f"Clave de Respuestas — {exam_name}"
    ws["A1"].font      = Font(bold=True, size=13, color=C_HDR_FG, name="Arial")
    ws["A1"].fill      = PatternFill("solid", fgColor=C_HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 3 — question numbers
    lbl = ws.cell(row=3, column=1, value="Pregunta")
    lbl.font = Font(bold=True, name="Arial", size=10)
    lbl.fill = PatternFill("solid", fgColor="ECF0F1")
    lbl.alignment = Alignment(horizontal="center")
    for i in range(n):
        cell = ws.cell(row=3, column=i + 2, value=i + 1)
        cell.font      = Font(bold=True, name="Arial", size=10)
        cell.fill      = PatternFill("solid", fgColor="ECF0F1")
        cell.alignment = Alignment(horizontal="center")

    # Row 4 — answers
    lbl2 = ws.cell(row=4, column=1, value="Respuesta")
    lbl2.font      = Font(bold=True, name="Arial", size=10, color=C_HDR_FG)
    lbl2.fill      = PatternFill("solid", fgColor=C_ACCENT)
    lbl2.alignment = Alignment(horizontal="center")
    for i, ans in enumerate(answer_key):
        cell = ws.cell(row=4, column=i + 2, value=ans)
        cell.font      = Font(bold=True, name="Arial", size=11)
        cell.fill      = PatternFill("solid", fgColor=C_GREEN)
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 12
    for i in range(n):
        ws.column_dimensions[get_column_letter(i + 2)].width = 5


# ── Reader helpers ─────────────────────────────────────────────────────────────

def read_answer_key_from_excel(excel_path):
    """Return the answer key list from a previously exported Excel, or None."""
    from openpyxl import load_workbook
    wb = load_workbook(excel_path, data_only=True)
    if "Clave de Respuestas" not in wb.sheetnames:
        return None
    ws = wb["Clave de Respuestas"]
    for row in ws.iter_rows():
        if row[0].value == "Respuesta":
            return [cell.value for cell in row[1:] if cell.value is not None]
    return None


def read_session_from_excel(excel_path):
    """
    Read a previously exported session.
    Returns (exam_name, answer_key, rows) where each row is a dict with:
        page_num, folio, grado, grupo, score, total, percentage,
        sk_average, confidence, error, mc_answers, sk_answers
    """
    from openpyxl import load_workbook
    wb = load_workbook(excel_path, data_only=True)

    # Exam name
    exam_name = "Examen"
    if "Resumen" in wb.sheetnames:
        title = wb["Resumen"]["A1"].value or ""
        if "—" in str(title):
            exam_name = str(title).split("—", 1)[1].strip()

    answer_key = read_answer_key_from_excel(excel_path) or []
    n_q        = len(answer_key)

    # Resumen sheet → base student data (header-based to handle Nombre column presence)
    resumen = {}
    if "Resumen" in wb.sheetnames:
        ws_res  = wb["Resumen"]
        hdr_row = [cell.value for cell in ws_res[4]]
        col_idx = {str(h): i for i, h in enumerate(hdr_row) if h is not None}
        def _rc(row, key, default=None):
            i = col_idx.get(key)
            return row[i] if i is not None and i < len(row) else default
        for row in ws_res.iter_rows(min_row=5, values_only=True):
            if row[0] is None:
                break
            folio = str(row[0])
            grado_val  = _rc(row, "Grado")
            grupo_val  = _rc(row, "Grupo")
            nombre_val = _rc(row, "Nombre")
            score_val  = _rc(row, "Puntaje", 0)
            total_val  = _rc(row, "Total", n_q)
            pct_val    = _rc(row, "Porcentaje (%)", 0.0)
            sk_val     = _rc(row, "Prom. Autoconoc.")
            conf_val   = _rc(row, "Confianza (%)", 0)
            err_val    = _rc(row, "Estado")
            resumen[folio] = {
                "folio":      folio,
                "nombre":     str(nombre_val).strip() if nombre_val else "",
                "grado":      str(grado_val) if grado_val else None,
                "grupo":      str(grupo_val) if grupo_val else None,
                "score":      score_val or 0,
                "total":      total_val or n_q,
                "percentage": pct_val or 0.0,
                "sk_average": sk_val if sk_val not in (None, "N/A") else None,
                "confidence": (conf_val or 0) / 100,
                "error":      str(err_val) if err_val and str(err_val) != "OK" else None,
            }

    # Detalle Preguntas → page_num + mc_answers
    detail = {}
    if "Detalle Preguntas" in wb.sheetnames:
        for row in wb["Detalle Preguntas"].iter_rows(min_row=4, values_only=True):
            if row[0] is None or row[1] is None:
                break
            folio = str(row[1])
            detail[folio] = {
                "page_num":   int(row[0]) if row[0] else 0,
                "mc_answers": [v if v != "-" else None
                               for v in row[4:4 + n_q]],
            }

    # Autoconocimiento → sk_answers (header-based to handle optional Nombre column)
    sk_data = {}
    if "Autoconocimiento" in wb.sheetnames:
        ws_sk   = wb["Autoconocimiento"]
        sk_hdrs = [cell.value for cell in ws_sk[2]]
        sk1_col = next((i for i, h in enumerate(sk_hdrs) if h == "SK1"), 3)
        for row in ws_sk.iter_rows(min_row=3, values_only=True):
            if row[0] is None:
                break
            folio = str(row[0])
            sk_data[folio] = [
                v if v != "-" else None
                for v in row[sk1_col:sk1_col + 10]
            ]

    # Merge and sort
    rows = []
    for folio, rd in resumen.items():
        dd  = detail.get(folio, {})
        row = {**rd,
               "page_num":   dd.get("page_num", 0),
               "mc_answers": dd.get("mc_answers", [None] * n_q),
               "sk_answers": sk_data.get(folio, [None] * 10),
               "nombre":     rd.get("nombre", "")}
        rows.append(row)
    rows.sort(key=lambda r: r["page_num"])
    return exam_name, answer_key, rows
