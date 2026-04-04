"""
Smart Test Assembly GUI
-----------------------
Tkinter-based GUI with four workflow tabs:
  1. Bank – load and inspect question bank
  2. Configure – assembly settings (bins, pairs, scoring weights)
  3. Generate – run assembly + analysis
  4. Results – summary table + status log

Design principles
-----------------
* No horizontal/vertical scrolling caused by oversized widgets.
* Grid-based layouts for bins and pair distribution tables.
* All blocking operations (file I/O, assembly, analysis) run in a
  background thread so the GUI stays responsive.
* Validation feedback is immediate and inline.
"""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from assembly_engine import AssemblyConfig, AssemblyEngine, build_config_from_gui
from data_loader import load_bank, get_pair_summary
from output_generator import generate_all_outputs
from psychometric import SimulationConfig, analyse_all_forms

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Colour palette & theme constants
# ---------------------------------------------------------------------------
BG_DARK    = "#1F3864"
BG_MID     = "#2E75B6"
BG_LIGHT   = "#EBF3FB"
ACCENT     = "#ED7D31"
TEXT_DARK  = "#1A1A2E"
TEXT_LIGHT = "#FFFFFF"
ENTRY_BG   = "#FFFFFF"
ENTRY_FG   = "#000000"
BTN_ACTIVE = "#C55A11"

FONT_TITLE  = ("Segoe UI", 15, "bold")
FONT_HDR    = ("Segoe UI", 11, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_MONO   = ("Consolas", 9)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _entry(parent, width=14, **kw) -> tk.Entry:
    e = tk.Entry(parent, width=width, bg=ENTRY_BG, fg=ENTRY_FG,
                 font=FONT_BODY, relief="solid", bd=1, **kw)
    return e


def _label(parent, text, bold=False, **kw) -> tk.Label:
    font = FONT_HDR if bold else FONT_BODY
    return tk.Label(parent, text=text, bg=parent.cget("bg"),
                    fg=TEXT_DARK, font=font, **kw)


def _btn(parent, text, command, bg=BG_MID, **kw) -> tk.Button:
    b = tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=TEXT_LIGHT, font=FONT_BODY,
        activebackground=BTN_ACTIVE, activeforeground=TEXT_LIGHT,
        relief="flat", padx=10, pady=5, cursor="hand2", **kw,
    )
    return b


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Test Assembly & Psychometric Analysis")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)
        self.minsize(900, 660)

        self.bank_df: Optional[pd.DataFrame] = None
        self.assembly_result = None
        self.analyses = None
        self.output_paths: dict[str, Path] = {}

        self._msg_queue: queue.Queue = queue.Queue()

        self._build_header()
        self._build_notebook()
        self._build_statusbar()

        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build_header(self):
        frm = tk.Frame(self, bg=BG_DARK, pady=8)
        frm.pack(fill="x")
        tk.Label(
            frm,
            text="Smart Test Assembly  &  Psychometric Analysis",
            bg=BG_DARK, fg=TEXT_LIGHT, font=FONT_TITLE,
        ).pack()

    def _build_notebook(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background=BG_DARK, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background=BG_MID, foreground=TEXT_LIGHT,
            font=FONT_HDR, padding=[14, 6],
        )
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", TEXT_LIGHT)])

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        self.tab_bank    = self._make_scrollable_tab()
        self.tab_config  = self._make_scrollable_tab()
        self.tab_generate = tk.Frame(self.nb, bg=BG_LIGHT)
        self.tab_results = self._make_scrollable_tab()

        self.nb.add(self.tab_bank,     text="  1. Bank  ")
        self.nb.add(self.tab_config,   text="  2. Configure  ")
        self.nb.add(self.tab_generate, text="  3. Generate  ")
        self.nb.add(self.tab_results,  text="  4. Results  ")

        self._build_bank_tab()
        self._build_config_tab()
        self._build_generate_tab()
        self._build_results_tab()

    def _make_scrollable_tab(self) -> tk.Frame:
        """Returns a plain Frame with a scrollable interior."""
        outer = tk.Frame(self.nb, bg=BG_LIGHT)
        canvas = tk.Canvas(outer, bg=BG_LIGHT, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=BG_LIGHT)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        def _on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_resize(event):
            canvas.itemconfig(window_id, width=event.width)
        inner.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_canvas_resize)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(
            int(-1 * (e.delta / 120)), "units"
        ))
        # store reference so tabs can add children to inner
        outer._inner = inner
        return outer

    def _build_statusbar(self):
        frm = tk.Frame(self, bg=BG_DARK, height=24)
        frm.pack(fill="x", side="bottom")
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(
            frm, textvariable=self._status_var,
            bg=BG_DARK, fg=TEXT_LIGHT, font=FONT_MONO,
            anchor="w", padx=8,
        ).pack(fill="x")

    # ------------------------------------------------------------------
    # Tab 1: Bank
    # ------------------------------------------------------------------

    def _build_bank_tab(self):
        p = self.tab_bank._inner
        p.configure(bg=BG_LIGHT, padx=20, pady=16)

        _label(p, "Question Bank", bold=True).grid(row=0, column=0, sticky="w", pady=(0, 4))

        frm_load = tk.Frame(p, bg=BG_LIGHT)
        frm_load.grid(row=1, column=0, sticky="ew", pady=4)

        self._bank_path_var = tk.StringVar(value="No file selected")
        tk.Label(
            frm_load, textvariable=self._bank_path_var,
            bg=BG_LIGHT, fg=BG_DARK, font=FONT_BODY,
            width=60, anchor="w",
        ).pack(side="left", padx=(0, 8))
        _btn(frm_load, "Browse…", self._browse_bank).pack(side="left")

        # Bank stats
        self._bank_stats_var = tk.StringVar(value="")
        tk.Label(
            p, textvariable=self._bank_stats_var,
            bg=BG_LIGHT, fg=BG_MID, font=FONT_BODY, justify="left",
        ).grid(row=2, column=0, sticky="w", pady=6)

        # Pair summary table
        _label(p, "(الناتج, المؤشر) Pair Summary", bold=True).grid(
            row=3, column=0, sticky="w", pady=(12, 4)
        )
        frm_table = tk.Frame(p, bg=BG_LIGHT)
        frm_table.grid(row=4, column=0, sticky="nsew")
        self._pair_tree = self._build_treeview(
            frm_table,
            columns=["الناتج", "المؤشر", "Count", "Mean b", "Min b", "Max b"],
            col_widths=[180, 180, 70, 80, 80, 80],
        )

    def _browse_bank(self):
        path = fd.askopenfilename(
            title="Select Question Bank",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")],
        )
        if not path:
            return
        self._status("Loading bank…")
        threading.Thread(target=self._load_bank_thread, args=(path,), daemon=True).start()

    def _load_bank_thread(self, path: str):
        try:
            df = load_bank(path)
            self._msg_queue.put(("bank_loaded", df, path))
        except Exception as exc:
            self._msg_queue.put(("error", str(exc)))

    def _on_bank_loaded(self, df: pd.DataFrame, path: str):
        self.bank_df = df
        self._bank_path_var.set(Path(path).name)
        stats_text = (
            f"Loaded {len(df)} questions  |  "
            f"Pairs: {df['_pair_key'].nunique()}  |  "
            f"b range: [{df['b'].min():.2f}, {df['b'].max():.2f}]  |  "
            f"Mean b: {df['b'].mean():.3f}"
        )
        self._bank_stats_var.set(stats_text)

        # Populate pair summary tree
        summary = get_pair_summary(df)
        for item in self._pair_tree.get_children():
            self._pair_tree.delete(item)
        for _, row in summary.iterrows():
            self._pair_tree.insert(
                "", "end",
                values=[
                    row["الناتج"], row["المؤشر"],
                    row["count"],
                    f"{row['mean_b']:.3f}", f"{row['min_b']:.3f}", f"{row['max_b']:.3f}",
                ],
            )

        # Pre-populate pair requirements in Configure tab
        self._populate_pair_requirements(df)
        self._status(f"Bank loaded: {len(df)} questions, {df['_pair_key'].nunique()} pairs.")

    # ------------------------------------------------------------------
    # Tab 2: Configure
    # ------------------------------------------------------------------

    def _build_config_tab(self):
        p = self.tab_config._inner
        p.configure(bg=BG_LIGHT, padx=20, pady=16)
        p.columnconfigure(0, weight=1)
        p.columnconfigure(1, weight=1)

        row = 0
        _label(p, "Assembly Settings", bold=True).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )
        row += 1

        # Basic params
        params_frm = tk.LabelFrame(
            p, text="Basic Parameters", bg=BG_LIGHT,
            fg=BG_DARK, font=FONT_HDR, padx=10, pady=8
        )
        params_frm.grid(row=row, column=0, sticky="nsew", padx=(0, 8), pady=4)
        row += 1

        self._cfg_n_forms    = self._labeled_entry(params_frm, "Number of Forms:", 0, default="4")
        self._cfg_q_per_form = self._labeled_entry(params_frm, "Questions / Form:", 1, default="40")
        self._cfg_mean_b     = self._labeled_entry(params_frm, "Target Mean b:", 2, default="0.0")
        self._cfg_min_b      = self._labeled_entry(params_frm, "Min b:", 3, default="-3.0")
        self._cfg_max_b      = self._labeled_entry(params_frm, "Max b:", 4, default="3.0")
        self._cfg_seed       = self._labeled_entry(params_frm, "RNG Seed (blank=random):", 5, default="")

        self._cfg_reuse_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            params_frm, text="Allow reuse across forms",
            variable=self._cfg_reuse_var, bg=BG_LIGHT, fg=TEXT_DARK, font=FONT_BODY,
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=4)

        # Scoring weights
        weights_frm = tk.LabelFrame(
            p, text="Scoring Weights (must sum ~1.0)", bg=BG_LIGHT,
            fg=BG_DARK, font=FONT_HDR, padx=10, pady=8
        )
        weights_frm.grid(row=row - 1, column=1, sticky="nsew", padx=(8, 0), pady=4)

        self._w_difficulty = self._labeled_entry(weights_frm, "Difficulty Fit:", 0, default="0.40")
        self._w_bin        = self._labeled_entry(weights_frm, "Bin Fit:", 1, default="0.25")
        self._w_usage      = self._labeled_entry(weights_frm, "Usage Penalty:", 2, default="0.20")
        self._w_random     = self._labeled_entry(weights_frm, "Randomness:", 3, default="0.15")

        # Simulation settings
        sim_frm = tk.LabelFrame(
            p, text="Simulation Settings", bg=BG_LIGHT,
            fg=BG_DARK, font=FONT_HDR, padx=10, pady=8
        )
        sim_frm.grid(row=row, column=0, sticky="nsew", padx=(0, 8), pady=4)
        row += 1

        self._sim_n         = self._labeled_entry(sim_frm, "Number of students:", 0, default="3000")
        self._sim_theta_mean = self._labeled_entry(sim_frm, "θ mean:", 1, default="0.0")
        self._sim_theta_sd   = self._labeled_entry(sim_frm, "θ SD:", 2, default="1.0")

        # Output directory
        out_frm = tk.LabelFrame(
            p, text="Output Directory", bg=BG_LIGHT,
            fg=BG_DARK, font=FONT_HDR, padx=10, pady=8
        )
        out_frm.grid(row=row - 1, column=1, sticky="nsew", padx=(8, 0), pady=4)

        self._output_dir_var = tk.StringVar(value=str(Path.home() / "TestOutput"))
        tk.Entry(
            out_frm, textvariable=self._output_dir_var,
            width=32, font=FONT_BODY, bg=ENTRY_BG, fg=ENTRY_FG,
        ).grid(row=0, column=0, sticky="ew", pady=4)
        _btn(out_frm, "Browse…", self._browse_output_dir).grid(row=0, column=1, padx=(6, 0))
        out_frm.columnconfigure(0, weight=1)

        # Prefix
        self._output_prefix = self._labeled_entry(out_frm, "File Prefix:", 1, default="test")

        # Bins section
        bins_frm = tk.LabelFrame(
            p, text="Difficulty Bins (proportion must sum to 1.0)",
            bg=BG_LIGHT, fg=BG_DARK, font=FONT_HDR, padx=10, pady=8,
        )
        bins_frm.grid(row=row, column=0, columnspan=2, sticky="ew", pady=6)
        row += 1
        self._build_bins_section(bins_frm)

        # Pair requirements
        pair_frm = tk.LabelFrame(
            p, text="(الناتج , المؤشر) Pair Requirements",
            bg=BG_LIGHT, fg=BG_DARK, font=FONT_HDR, padx=10, pady=8,
        )
        pair_frm.grid(row=row, column=0, columnspan=2, sticky="ew", pady=6)
        row += 1
        self._build_pair_requirements_section(pair_frm)

    def _labeled_entry(self, parent, label: str, row: int, default: str = "") -> tk.Entry:
        tk.Label(
            parent, text=label, bg=parent.cget("bg"),
            fg=TEXT_DARK, font=FONT_BODY, anchor="w",
        ).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        e = _entry(parent, width=14)
        e.insert(0, default)
        e.grid(row=row, column=1, sticky="w", pady=2)
        return e

    def _build_bins_section(self, parent):
        hdr_row = tk.Frame(parent, bg=BG_LIGHT)
        hdr_row.pack(fill="x", pady=(0, 4))
        for i, (text, w) in enumerate([("Low b", 8), ("High b", 8), ("Proportion", 10)]):
            tk.Label(
                hdr_row, text=text, bg=BG_MID, fg=TEXT_LIGHT,
                font=FONT_HDR, width=w, relief="flat", padx=4,
            ).grid(row=0, column=i, padx=2)
        tk.Label(hdr_row, text="", bg=BG_LIGHT, width=6).grid(row=0, column=3)

        self._bins_frame = tk.Frame(parent, bg=BG_LIGHT)
        self._bins_frame.pack(fill="x")
        self._bin_rows: list[tuple[tk.Entry, tk.Entry, tk.Entry]] = []

        ctrl_row = tk.Frame(parent, bg=BG_LIGHT)
        ctrl_row.pack(fill="x", pady=(6, 0))
        _btn(ctrl_row, "+ Add Bin", self._add_bin_row, bg=BG_MID).pack(side="left", padx=4)
        _btn(ctrl_row, "Clear Bins", self._clear_bins, bg="#C00000").pack(side="left", padx=4)
        _btn(ctrl_row, "Auto 3 Bins", lambda: self._auto_bins(3), bg=BG_DARK).pack(side="left", padx=4)

        # Default 3 bins
        defaults = [(-3.0, -0.5, 0.25), (-0.5, 0.5, 0.50), (0.5, 3.0, 0.25)]
        for lo, hi, prop in defaults:
            self._add_bin_row(lo=lo, hi=hi, prop=prop)

    def _add_bin_row(self, lo: float = -1.0, hi: float = 1.0, prop: float = 0.33):
        frm = tk.Frame(self._bins_frame, bg=BG_LIGHT)
        frm.pack(fill="x", pady=2)
        e_lo   = _entry(frm, width=10); e_lo.insert(0, str(lo))
        e_hi   = _entry(frm, width=10); e_hi.insert(0, str(hi))
        e_prop = _entry(frm, width=10); e_prop.insert(0, str(prop))
        e_lo.grid(row=0, column=0, padx=4)
        e_hi.grid(row=0, column=1, padx=4)
        e_prop.grid(row=0, column=2, padx=4)
        row_tuple = (e_lo, e_hi, e_prop)
        self._bin_rows.append(row_tuple)
        del_btn = _btn(
            frm, "×", lambda rt=row_tuple, f=frm: self._remove_bin_row(rt, f),
            bg="#C00000",
        )
        del_btn.config(width=2, pady=2)
        del_btn.grid(row=0, column=3, padx=4)

    def _remove_bin_row(self, row_tuple, frm):
        if row_tuple in self._bin_rows:
            self._bin_rows.remove(row_tuple)
        frm.destroy()

    def _clear_bins(self):
        for widget in self._bins_frame.winfo_children():
            widget.destroy()
        self._bin_rows.clear()

    def _auto_bins(self, n: int):
        self._clear_bins()
        min_b = float(self._cfg_min_b.get() or -3.0)
        max_b = float(self._cfg_max_b.get() or 3.0)
        step = (max_b - min_b) / n
        prop = round(1.0 / n, 4)
        for i in range(n):
            lo = round(min_b + i * step, 3)
            hi = round(min_b + (i + 1) * step, 3)
            self._add_bin_row(lo=lo, hi=hi, prop=prop)

    def _build_pair_requirements_section(self, parent):
        note = tk.Label(
            parent,
            text="Load a bank first to auto-populate. Adjust counts as needed.",
            bg=BG_LIGHT, fg=BG_MID, font=FONT_BODY,
        )
        note.pack(anchor="w", pady=(0, 4))

        # Header
        hdr = tk.Frame(parent, bg=BG_LIGHT)
        hdr.pack(fill="x", pady=(0, 2))
        for i, (text, w) in enumerate([
            ("الناتج", 20), ("المؤشر", 20), ("Count", 8), ("Available", 10)
        ]):
            tk.Label(
                hdr, text=text, bg=BG_MID, fg=TEXT_LIGHT,
                font=FONT_HDR, width=w, relief="flat", padx=4,
            ).grid(row=0, column=i, padx=2)

        self._pair_req_frame = tk.Frame(parent, bg=BG_LIGHT)
        self._pair_req_frame.pack(fill="x")
        self._pair_req_rows: list[dict] = []

        # Summary label
        self._pair_total_label = tk.Label(
            parent, text="", bg=BG_LIGHT, fg=BG_DARK, font=FONT_BODY
        )
        self._pair_total_label.pack(anchor="w", pady=4)

    def _populate_pair_requirements(self, df: pd.DataFrame):
        """Called after bank load to populate pair requirement rows."""
        for widget in self._pair_req_frame.winfo_children():
            widget.destroy()
        self._pair_req_rows.clear()

        from data_loader import get_pair_summary
        summary = get_pair_summary(df)

        n_pairs = len(summary)
        if n_pairs == 0:
            return

        try:
            q_per_form = int(self._cfg_q_per_form.get())
        except ValueError:
            q_per_form = 40

        counts = self._distribute_equal(q_per_form, n_pairs)

        for idx, (_, row) in enumerate(summary.iterrows()):
            frm = tk.Frame(self._pair_req_frame, bg=BG_LIGHT)
            frm.pack(fill="x", pady=1)

            e_natej   = tk.Entry(frm, width=22, bg="#F0F0F0", fg=TEXT_DARK, font=FONT_BODY, state="readonly")
            e_moshir  = tk.Entry(frm, width=22, bg="#F0F0F0", fg=TEXT_DARK, font=FONT_BODY, state="readonly")
            e_count   = _entry(frm, width=8)
            e_avail   = tk.Entry(frm, width=10, bg="#F0F0F0", fg=TEXT_DARK, font=FONT_BODY, state="readonly")

            e_natej.config(state="normal"); e_natej.insert(0, str(row["الناتج"])); e_natej.config(state="readonly")
            e_moshir.config(state="normal"); e_moshir.insert(0, str(row["المؤشر"])); e_moshir.config(state="readonly")
            e_count.insert(0, str(counts[idx]))
            e_avail.config(state="normal"); e_avail.insert(0, str(int(row["count"]))); e_avail.config(state="readonly")

            e_natej.grid(row=0, column=0, padx=4)
            e_moshir.grid(row=0, column=1, padx=4)
            e_count.grid(row=0, column=2, padx=4)
            e_avail.grid(row=0, column=3, padx=4)

            pair_key = str(row["الناتج"]) + "||" + str(row["المؤشر"])
            self._pair_req_rows.append({
                "pair_key": pair_key,
                "e_natej": e_natej,
                "e_moshir": e_moshir,
                "e_count": e_count,
                "e_avail": e_avail,
            })

        self._update_pair_total()
        self._pair_req_frame.bind_all("<<Modified>>", lambda e: self._update_pair_total())

    def _update_pair_total(self):
        total = 0
        for row in self._pair_req_rows:
            try:
                total += int(row["e_count"].get())
            except ValueError:
                pass
        try:
            q_per_form = int(self._cfg_q_per_form.get())
        except ValueError:
            q_per_form = 40
        color = "green" if total == q_per_form else "#C00000"
        self._pair_total_label.config(
            text=f"Total assigned: {total}  (target: {q_per_form})",
            fg=color,
        )

    @staticmethod
    def _distribute_equal(total: int, n: int) -> list[int]:
        base = total // n
        remainder = total % n
        counts = [base] * n
        for i in range(remainder):
            counts[i] += 1
        return counts

    def _browse_output_dir(self):
        d = fd.askdirectory(title="Select Output Directory")
        if d:
            self._output_dir_var.set(d)

    # ------------------------------------------------------------------
    # Tab 3: Generate
    # ------------------------------------------------------------------

    def _build_generate_tab(self):
        p = self.tab_generate
        p.configure(bg=BG_LIGHT)

        top = tk.Frame(p, bg=BG_LIGHT, pady=14)
        top.pack(fill="x", padx=20)
        _label(top, "Generate Forms & Run Psychometric Analysis", bold=True).pack(anchor="w")

        btn_frm = tk.Frame(p, bg=BG_LIGHT, pady=6)
        btn_frm.pack(padx=20, anchor="w")
        self._gen_btn = _btn(
            btn_frm, "▶  Generate & Analyse",
            self._run_generation, bg=ACCENT,
        )
        self._gen_btn.config(font=("Segoe UI", 11, "bold"), pady=8, padx=20)
        self._gen_btn.pack(side="left", padx=(0, 12))

        self._progress_var = tk.DoubleVar(value=0)
        self._progress = ttk.Progressbar(
            btn_frm, variable=self._progress_var,
            maximum=100, length=300, mode="determinate",
        )
        self._progress.pack(side="left")

        # Log area
        log_frm = tk.LabelFrame(
            p, text="Log", bg=BG_LIGHT, fg=BG_DARK,
            font=FONT_HDR, padx=8, pady=8,
        )
        log_frm.pack(fill="both", expand=True, padx=20, pady=8)

        self._log_text = tk.Text(
            log_frm, height=20, bg="#0D1117", fg="#58D68D",
            font=FONT_MONO, state="disabled", relief="flat",
        )
        vsb = ttk.Scrollbar(log_frm, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._log_text.pack(fill="both", expand=True)

    def _run_generation(self):
        if self.bank_df is None:
            mb.showerror("No Bank", "Please load a question bank first (Tab 1).")
            return

        try:
            params = self._collect_params()
        except ValueError as exc:
            mb.showerror("Configuration Error", str(exc))
            return

        self._gen_btn.config(state="disabled")
        self._progress_var.set(0)
        self._log_clear()

        threading.Thread(
            target=self._generation_thread,
            args=(params,),
            daemon=True,
        ).start()

    def _generation_thread(self, params: dict):
        try:
            self._log("Building assembly configuration…")
            config = build_config_from_gui(params)

            self._log(f"Starting assembly: {config.n_forms} forms × {config.questions_per_form} questions")
            engine = AssemblyEngine(self.bank_df, config)
            result = engine.assemble()
            self.assembly_result = result
            self._msg_queue.put(("progress", 40))

            if result.warnings:
                for w in result.warnings:
                    self._log(f"  ⚠ {w}", color="yellow")

            form_names = [f"Form {i + 1}" for i in range(len(result.forms))]
            self._log("Running psychometric analysis…")

            sim_cfg = SimulationConfig(
                n_students=params.get("n_students", 3000),
                theta_mean=params.get("theta_mean", 0.0),
                theta_sd=params.get("theta_sd", 1.0),
                rng_seed=params.get("rng_seed"),
            )
            analyses = analyse_all_forms(result.forms, form_names, sim_cfg=sim_cfg)
            self.analyses = analyses
            self._msg_queue.put(("progress", 75))

            self._log("Writing output files…")
            out_dir = Path(params["output_dir"])
            prefix  = params.get("prefix", "test")
            paths = generate_all_outputs(analyses, out_dir, prefix=prefix)
            self.output_paths = paths
            self._msg_queue.put(("progress", 100))

            summary_lines = ["\n── Output files ──"]
            for label, p in paths.items():
                summary_lines.append(f"  [{label}]  {p}")
            self._log("\n".join(summary_lines))
            self._log("\n✓ Done!", color="lime")
            self._msg_queue.put(("generation_done", analyses))

        except Exception as exc:
            self._log(f"\n✗ Error: {exc}", color="red")
            logger.exception("Generation failed")
            self._msg_queue.put(("generation_error",))

    def _collect_params(self) -> dict:
        def _float(e, label):
            v = e.get().strip()
            try:
                return float(v)
            except ValueError:
                raise ValueError(f"Invalid value for '{label}': '{v}'")

        def _int(e, label):
            v = e.get().strip()
            try:
                return int(v)
            except ValueError:
                raise ValueError(f"Invalid value for '{label}': '{v}'")

        n_forms    = _int(self._cfg_n_forms, "Number of Forms")
        q_per_form = _int(self._cfg_q_per_form, "Questions / Form")
        mean_b     = _float(self._cfg_mean_b, "Target Mean b")
        min_b      = _float(self._cfg_min_b, "Min b")
        max_b      = _float(self._cfg_max_b, "Max b")

        seed_str = self._cfg_seed.get().strip()
        seed = int(seed_str) if seed_str else None

        # Bins
        bins = []
        for e_lo, e_hi, e_prop in self._bin_rows:
            try:
                lo   = float(e_lo.get())
                hi   = float(e_hi.get())
                prop = float(e_prop.get())
                bins.append({"low": lo, "high": hi, "proportion": prop})
            except ValueError as exc:
                raise ValueError(f"Invalid bin value: {exc}")

        if bins:
            total_prop = sum(b["proportion"] for b in bins)
            if not (0.98 <= total_prop <= 1.02):
                raise ValueError(
                    f"Bin proportions must sum to 1.0 (currently {total_prop:.4f})."
                )

        # Pair requirements
        pair_reqs = []
        for row in self._pair_req_rows:
            try:
                count = int(row["e_count"].get())
            except ValueError:
                count = 0
            pair_reqs.append({
                "pair_key": row["pair_key"],
                "الناتج":  row["e_natej"].get(),
                "المؤشر":  row["e_moshir"].get(),
                "count":   count,
            })

        if pair_reqs:
            total = sum(r["count"] for r in pair_reqs)
            if total != q_per_form:
                raise ValueError(
                    f"Pair requirement total ({total}) must equal Questions/Form ({q_per_form})."
                )

        # Weights
        w_diff   = _float(self._w_difficulty, "Weight: Difficulty Fit")
        w_bin    = _float(self._w_bin, "Weight: Bin Fit")
        w_usage  = _float(self._w_usage, "Weight: Usage Penalty")
        w_random = _float(self._w_random, "Weight: Randomness")

        # Simulation
        n_students  = _int(self._sim_n, "Number of students")
        theta_mean  = _float(self._sim_theta_mean, "θ mean")
        theta_sd    = _float(self._sim_theta_sd, "θ SD")

        return {
            "n_forms": n_forms,
            "questions_per_form": q_per_form,
            "target_mean_b": mean_b,
            "target_min_b": min_b,
            "target_max_b": max_b,
            "bins": bins,
            "pair_requirements": pair_reqs,
            "w_difficulty": w_diff,
            "w_bin": w_bin,
            "w_usage": w_usage,
            "w_random": w_random,
            "allow_reuse_across_forms": self._cfg_reuse_var.get(),
            "rng_seed": seed,
            "output_dir": self._output_dir_var.get(),
            "prefix": self._cfg_prefix_var(),
            "n_students": n_students,
            "theta_mean": theta_mean,
            "theta_sd": theta_sd,
        }

    def _cfg_prefix_var(self) -> str:
        return getattr(self, "_output_prefix", None) and self._output_prefix.get() or "test"

    # ------------------------------------------------------------------
    # Tab 4: Results
    # ------------------------------------------------------------------

    def _build_results_tab(self):
        p = self.tab_results._inner
        p.configure(bg=BG_LIGHT, padx=20, pady=16)
        p.columnconfigure(0, weight=1)

        _label(p, "Results Summary", bold=True).grid(row=0, column=0, sticky="w", pady=(0, 8))

        self._results_tree = self._build_treeview(
            p,
            columns=[
                "Form", "N Items", "Mean b", "SD b", "Min b", "Max b",
                "Mean a", "Mean c", "Alpha", "Mean Corr.",
            ],
            col_widths=[90, 70, 80, 70, 70, 70, 70, 70, 70, 90],
            row=1,
        )

        # Open output folder button
        btn_frm = tk.Frame(p, bg=BG_LIGHT)
        btn_frm.grid(row=2, column=0, sticky="w", pady=8)
        _btn(btn_frm, "Open Output Folder", self._open_output_folder).pack(side="left")
        _btn(btn_frm, "Refresh", self._refresh_results, bg=BG_DARK).pack(side="left", padx=6)

    def _refresh_results(self):
        if self.analyses is None:
            return
        for item in self._results_tree.get_children():
            self._results_tree.delete(item)
        for fa in self.analyses:
            s = fa.summary
            self._results_tree.insert(
                "", "end",
                values=[
                    fa.form_name,
                    s.get("n_items", ""),
                    f"{s.get('mean_b', 0):.3f}",
                    f"{s.get('std_b', 0):.3f}",
                    f"{s.get('min_b', 0):.3f}",
                    f"{s.get('max_b', 0):.3f}",
                    f"{s.get('mean_a', 0):.3f}",
                    f"{s.get('mean_c', 0):.3f}",
                    f"{s.get('alpha', 0):.3f}",
                    f"{s.get('mean_item_cor', 0):.3f}",
                ],
            )

    def _open_output_folder(self):
        import subprocess, sys
        out_dir = self._output_dir_var.get()
        if sys.platform == "win32":
            subprocess.Popen(["explorer", out_dir])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", out_dir])
        else:
            subprocess.Popen(["xdg-open", out_dir])

    # ------------------------------------------------------------------
    # Shared Treeview builder
    # ------------------------------------------------------------------

    def _build_treeview(
        self,
        parent,
        columns: list[str],
        col_widths: list[int],
        row: int = 0,
    ) -> ttk.Treeview:
        style = ttk.Style()
        style.configure(
            "Custom.Treeview.Heading",
            background=BG_MID, foreground=TEXT_LIGHT, font=FONT_HDR,
        )
        style.configure("Custom.Treeview", font=FONT_BODY, rowheight=22)
        style.map("Custom.Treeview", background=[("selected", ACCENT)])

        frm = tk.Frame(parent, bg=BG_LIGHT)
        if isinstance(parent, ttk.Notebook):
            frm.grid(row=row, column=0, sticky="nsew")
        elif hasattr(parent, "_inner"):
            frm.pack(fill="both", expand=True)
        else:
            frm.grid(row=row, column=0, sticky="nsew")

        tree = ttk.Treeview(
            frm, columns=columns, show="headings",
            style="Custom.Treeview",
        )
        vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col, w in zip(columns, col_widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, minwidth=50, anchor="center")

        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree.grid(row=0, column=0, sticky="nsew")
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)
        return tree

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------

    def _log(self, msg: str, color: str = ""):
        self._msg_queue.put(("log", msg, color))

    def _log_clear(self):
        self._msg_queue.put(("log_clear",))

    def _append_log(self, msg: str, color: str = ""):
        self._log_text.config(state="normal")
        tag = color if color else "default"
        self._log_text.tag_config(tag, foreground=color or "#58D68D")
        self._log_text.insert("end", msg + "\n", tag)
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    # ------------------------------------------------------------------
    # Message queue processing (runs in main thread)
    # ------------------------------------------------------------------

    def _poll_queue(self):
        try:
            while True:
                item = self._msg_queue.get_nowait()
                kind = item[0]

                if kind == "bank_loaded":
                    _, df, path = item
                    self._on_bank_loaded(df, path)

                elif kind == "log":
                    _, msg = item[0], item[1]
                    color = item[2] if len(item) > 2 else ""
                    self._append_log(msg, color)

                elif kind == "log_clear":
                    self._log_text.config(state="normal")
                    self._log_text.delete("1.0", "end")
                    self._log_text.config(state="disabled")

                elif kind == "progress":
                    self._progress_var.set(item[1])

                elif kind == "generation_done":
                    self.analyses = item[1]
                    self._gen_btn.config(state="normal")
                    self._progress_var.set(100)
                    self._refresh_results()
                    self.nb.select(3)  # jump to Results tab
                    self._status("Generation complete.")
                    mb.showinfo(
                        "Done",
                        "All forms generated and analysed.\nOutput files written successfully.",
                    )

                elif kind == "generation_error":
                    self._gen_btn.config(state="normal")
                    self._status("Generation failed – see log.")

                elif kind == "error":
                    self._status(f"Error: {item[1]}")
                    mb.showerror("Error", item[1])

        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

    def _status(self, msg: str):
        self._status_var.set(msg)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def launch():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    )
    app = App()
    app.mainloop()


if __name__ == "__main__":
    launch()
