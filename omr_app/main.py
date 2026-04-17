"""
OMR Grader — Desktop GUI
"""
import os, sys, json, threading, tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime

from sheet_generator import generate_omr_sheet
from omr_scanner     import scan_pdf
from omr_grader      import grade_results
from omr_exporter    import export_to_excel

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "omr_config.json")

def load_config():
    d = {"exam_name":"Examen","num_mc_questions":20,
         "answer_key":[],
         "last_output_dir":os.path.expanduser("~")}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                d.update(json.load(f))
        except Exception:
            pass
    return d

def save_config(cfg):
    with open(CONFIG_FILE,"w") as f:
        json.dump(cfg, f, indent=2)

BG=    "#1E2B3C"
BG2=   "#253447"
ACCENT="#3498DB"
SUCCESS="#2ECC71"
ERROR= "#E74C3C"
FG=    "#ECF0F1"
FG2=   "#BDC3C7"
CARD=  "#2C3E50"


class OMRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.title("OMR Grader")
        self.geometry("880x720")
        self.configure(bg=BG)
        self.resizable(True, True)

        self.pdf_path          = tk.StringVar()
        self.exam_name_var     = tk.StringVar(value=self.cfg["exam_name"])
        self.num_questions_var = tk.IntVar(value=self.cfg["num_mc_questions"])
        self.status_var        = tk.StringVar(value="Listo.")
        self.progress_var      = tk.DoubleVar(value=0)

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋 OMR Grader", bg=ACCENT, fg="white",
                 font=("Arial",18,"bold")).pack(side="left", padx=20, pady=10)
        tk.Label(hdr, text="Calificación automática de exámenes",
                 bg=ACCENT, fg="#D6EAF8",
                 font=("Arial",10)).pack(side="left", padx=5, pady=14)

        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=18, pady=14)

        left  = tk.Frame(main, bg=BG)
        left.pack(side="left", fill="both", expand=True, padx=(0,10))
        right = tk.Frame(main, bg=BG, width=240)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        self._card(left, "⚙  Configuración del Examen", self._build_settings)
        self._card(left, "📂  Archivo PDF Escaneado",   self._build_pdf_zone)
        self._card(left, "🗝  Clave de Respuestas",      self._build_answer_key)
        self._build_actions(left)
        self._build_progress(left)
        self._build_right_panel(right)

    def _card(self, parent, title, builder_fn):
        frame = tk.Frame(parent, bg=CARD)
        frame.pack(fill="x", pady=(0,10))
        tk.Label(frame, text=title, bg=CARD, fg=ACCENT,
                 font=("Arial",10,"bold")).pack(anchor="w", padx=12, pady=(10,4))
        tk.Frame(frame, bg=ACCENT, height=1).pack(fill="x", padx=12)
        inner = tk.Frame(frame, bg=CARD)
        inner.pack(fill="x", padx=12, pady=8)
        builder_fn(inner)

    def _build_settings(self, f):
        # Row 1: exam name
        r1 = tk.Frame(f, bg=CARD); r1.pack(fill="x", pady=2)
        tk.Label(r1, text="Nombre del examen:", bg=CARD, fg=FG,
                 font=("Arial",9)).pack(side="left")
        tk.Entry(r1, textvariable=self.exam_name_var, bg=BG2, fg=FG,
                 insertbackground=FG, font=("Arial",10), bd=0,
                 width=26).pack(side="left", padx=(8,0))

        # Row 2: num questions
        r2 = tk.Frame(f, bg=CARD); r2.pack(fill="x", pady=6)
        tk.Label(r2, text="Nº preguntas Sec.1:", bg=CARD, fg=FG,
                 font=("Arial",9)).pack(side="left")
        self.spinbox = tk.Spinbox(r2, from_=5, to=40,
                                  textvariable=self.num_questions_var,
                                  width=4, bg=BG2, fg=FG,
                                  buttonbackground=BG2, font=("Arial",10),
                                  command=self._rebuild_answer_key)
        self.spinbox.pack(side="left", padx=(6,0))
        self.spinbox.bind("<FocusOut>", lambda e: self._safe_rebuild())
        self.spinbox.bind("<Return>",   lambda e: self._safe_rebuild())

    def _safe_rebuild(self):
        try:
            n = self.num_questions_var.get()
            if 1 <= n <= 40:
                self._rebuild_answer_key()
        except Exception:
            pass

    def _build_pdf_zone(self, f):
        self.drop_frame = tk.Frame(f, bg=BG2, bd=2, relief="solid",
                                   cursor="hand2")
        self.drop_frame.pack(fill="x", ipady=14)
        self.drop_label = tk.Label(self.drop_frame,
            text="Haz clic para seleccionar el PDF escaneado",
            bg=BG2, fg=FG2, font=("Arial",10))
        self.drop_label.pack(pady=10)
        self.drop_frame.bind("<Button-1>", lambda e: self._browse_pdf())
        self.drop_label.bind("<Button-1>", lambda e: self._browse_pdf())

    def _build_answer_key(self, f):
        self.key_frame = tk.Frame(f, bg=CARD)
        self.key_frame.pack(fill="x")
        self._rebuild_answer_key()

    def _rebuild_answer_key(self):
        for w in self.key_frame.winfo_children():
            w.destroy()
        n = self.num_questions_var.get()
        self.answer_vars = []
        cols = 5
        for i in range(n):
            row = i // cols; col = i % cols
            sub = tk.Frame(self.key_frame, bg=CARD)
            sub.grid(row=row, column=col, padx=4, pady=3)
            tk.Label(sub, text=f"{i+1}", bg=CARD, fg=FG2,
                     font=("Arial",7)).pack()
            var = tk.StringVar(value="A")
            if i < len(self.cfg.get("answer_key",[])):
                var.set(self.cfg["answer_key"][i])
            ttk.Combobox(sub, textvariable=var,
                         values=["A","B","C","D"],
                         width=3, state="readonly",
                         font=("Arial",9)).pack()
            self.answer_vars.append(var)

    def _build_actions(self, f):
        bf = tk.Frame(f, bg=BG); bf.pack(fill="x", pady=4)
        self._btn(bf, "🖨  Generar Hoja OMR",  self._generate_sheet, ACCENT)
        self._btn(bf, "▶  Calificar PDF",       self._start_grading,  SUCCESS)

    def _btn(self, parent, text, cmd, color):
        tk.Button(parent, text=text, command=cmd, bg=color, fg="white",
                  font=("Arial",10,"bold"), bd=0, padx=14, pady=8,
                  activebackground=color, cursor="hand2",
                  relief="flat").pack(side="left", padx=(0,8))

    def _build_progress(self, f):
        pf = tk.Frame(f, bg=BG); pf.pack(fill="x", pady=(6,0))
        ttk.Progressbar(pf, variable=self.progress_var,
                        maximum=100).pack(fill="x", pady=(0,4))
        tk.Label(pf, textvariable=self.status_var, bg=BG, fg=FG2,
                 font=("Arial",9)).pack(anchor="w")

    def _build_right_panel(self, f):
        tk.Label(f, text="📊 Registro", bg=BG, fg=ACCENT,
                 font=("Arial",11,"bold")).pack(anchor="w", padx=10, pady=(0,4))
        self.log_box = tk.Text(f, bg=CARD, fg=FG2, font=("Courier",8),
                               bd=0, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=8)
        sb = ttk.Scrollbar(f, command=self.log_box.yview)
        sb.pack(fill="y", side="right")
        self.log_box.configure(yscrollcommand=sb.set)
        tk.Label(f, text="─"*40, bg=BG, fg=BG2).pack()
        self._btn(f, "🗑  Limpiar", self._clear_log, "#555")

    # ── helpers ───────────────────────────────────────────────────────────────
    def _log(self, msg):
        self.log_box.configure(state="normal")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0","end")
        self.log_box.configure(state="disabled")

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Seleccionar PDF escaneado",
            filetypes=[("PDF","*.pdf")])
        if path:
            self.pdf_path.set(path)
            self.drop_label.configure(
                text=f"✅  {os.path.basename(path)}", fg=SUCCESS)
            self._log(f"PDF cargado: {os.path.basename(path)}")

    def _get_answer_key(self):
        return [v.get() for v in self.answer_vars]

    def _read_n(self):
        try:
            n = int(self.spinbox.get())
            self.num_questions_var.set(n)
            return n
        except Exception:
            return self.num_questions_var.get()

    # ── generate sheet ────────────────────────────────────────────────────────
    def _generate_sheet(self):
        n     = self._read_n()
        exam  = self.exam_name_var.get() or "Examen"
        odir  = self.cfg.get("last_output_dir", os.path.expanduser("~"))
        path  = filedialog.asksaveasfilename(
            title="Guardar hoja OMR", defaultextension=".pdf",
            initialdir=odir, filetypes=[("PDF","*.pdf")],
            initialfile=f"hoja_omr_{exam.replace(' ','_')}.pdf")
        if not path:
            return
        try:
            generate_omr_sheet(path, num_mc_questions=n, exam_name=exam)
            self._log(f"Hoja OMR generada: {os.path.basename(path)}")
            messagebox.showinfo("Éxito", f"Hoja OMR guardada en:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ── grading ───────────────────────────────────────────────────────────────
    def _start_grading(self):
        if not self.pdf_path.get():
            messagebox.showwarning("Aviso","Selecciona un PDF escaneado primero.")
            return
        answer_key = self._get_answer_key()
        exam  = self.exam_name_var.get() or "Examen"
        n     = self._read_n()
        odir  = self.cfg.get("last_output_dir", os.path.expanduser("~"))
        out   = filedialog.asksaveasfilename(
            title="Guardar resultados Excel", defaultextension=".xlsx",
            initialdir=odir, filetypes=[("Excel","*.xlsx")],
            initialfile=f"resultados_{exam.replace(' ','_')}.xlsx")
        if not out:
            return

        self.cfg.update({
            "last_output_dir": os.path.dirname(out),
            "exam_name": exam,
            "num_mc_questions": n,
            "answer_key": answer_key,
        })
        save_config(self.cfg)

        threading.Thread(
            target=self._run_grading,
            args=(self.pdf_path.get(), answer_key, exam, out, n),
            daemon=True).start()

    def _run_grading(self, pdf_path, answer_key, exam_name, out_path, n_questions):
        try:
            self.status_var.set("Procesando PDF...")
            self.progress_var.set(5)
            self._log(f"Iniciando escaneo: {os.path.basename(pdf_path)}")

            def cb(i, total):
                self.progress_var.set(10 + int(i/max(total,1)*70))
                self.status_var.set(f"Procesando página {i+1}/{total}...")
                self._log(f"  Página {i+1}/{total}")

            scans  = scan_pdf(pdf_path, n_questions, progress_callback=cb)
            self._log(f"Páginas procesadas: {len(scans)}")
            self.progress_var.set(80)

            self.status_var.set("Calificando...")
            graded = grade_results(scans, answer_key)
            for gr in graded:
                s = "✓" if gr.percentage >= 70 else "✗"
                self._log(f"  {s} Folio {gr.folio} | "
                          f"{gr.score}/{gr.total} ({gr.percentage}%)")

            self.progress_var.set(90)
            self.status_var.set("Exportando Excel...")
            export_to_excel(graded, answer_key, exam_name, out_path)

            self.progress_var.set(100)
            self.status_var.set(f"✅ Listo — {out_path}")
            self._log(f"Excel guardado: {os.path.basename(out_path)}")
            self.after(0, lambda: messagebox.showinfo(
                "Calificación Completa",
                f"Se calificaron {len(graded)} estudiantes.\n\n"
                f"Archivo guardado en:\n{out_path}"))
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            self._log(f"ERROR: {e}")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))


if __name__ == "__main__":
    OMRApp().mainloop()
