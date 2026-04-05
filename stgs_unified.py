"""
STGS — Smart Test Generation System
=====================================
Unified launcher with two completely separate assembly modes:

  MODE 1 — القدرات والتحصيل
    Standardised national exams (قدرات علمي، قدرات نظري، التحصيلي، …)
    Bank columns: Category, D, Difficulty, [تمييز, مستوى التمييز]
    Assembly by: Category → count
    No Subject/Grade filters

  MODE 2 — التقييم التربوي والمناهج
    Curriculum-based educational assessments (science, math, reading by grade)
    Bank columns: الناتج, المؤشر, المجال, difficulty, [subject, grade, language]
    Assembly by: الناتج alone  OR  (الناتج, المؤشر) pair
    Subject / Grade / Language filters available

Both modes share:
  • Scoring-based greedy selection (replaces random.sample + 100-retry)
  • Automatic 3PL psychometric analysis after assembly (3000 students)
  • Professional Excel output: Forms + 3PL Analysis + Remaining Questions

Dependencies: pandas numpy openpyxl matplotlib xlsxwriter
"""

from __future__ import annotations

import io, logging, math, os, queue, subprocess, sys, threading, tempfile
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import numpy as np
import pandas as pd
import xlsxwriter
from openpyxl import Workbook as OPWorkbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-7s  %(message)s")
log = logging.getLogger("stgs")


# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

# ── Mode 1: national/standardised exams ─────────────────────────────────────
MODE1_EXAMS: dict[str, dict[str, int]] = {
    "قدرات علمي": {
        "MAR": 5, "MAL": 1, "MAN": 2, "MGE": 3,
        "VAN": 3, "VCA": 2, "VSC": 3, "VRC": 5,
    },
    "قدرات نظري": {
        "حساب": 3, "بياني": 1, "هندسة": 2,
        "التناظر اللفظي": 5, "الخطأ السياقي": 4,
        "إكمال الجمل": 3, "استيعاب المقروء": 6,
    },
    "القدرة المعرفية": {
        "المترادفات والمتضادات": 2, "الاستدلال اللفظي": 2,
        "الاستيعاب اللفظي": 2, "العمليات الحسابية": 3,
        "سلاسل الأرقام": 2, "تفسير البيانات": 2,
        "سلاسل الأشكال": 1, "تطابق الأشكال": 1,
        "فتح الصندوق": 2, "الاستنباطي": 3,
    },
    "التحصيلي": {
        "أحياء": 18, "كيمياء": 18, "فيزياء": 18, "الرياضيات": 18,
    },
    "قدرات الجامعيين": {
        "حساب": 2, "جبر": 2, "بياني": 2, "هندسة": 2,
        "CT1": 2, "CT2": 3, "CT3": 5,
        "التناظر اللفظي": 2, "خطاء السياقي": 2,
        "إستيعاب المقروء": 2, "إكمال الجمل": 2,
    },
    "مخصص": {},
}

MODE1_STAGES: dict[str, dict] = {
    "Stage 1 — عام":   {"Range": (0.20, 0.90), "Mean": (0.52, 0.58)},
    "سهل (E)":          {"Range": (0.00, 0.45), "Mean": (0.27, 0.35)},
    "متوسط (M)":        {"Range": (0.40, 0.75), "Mean": (0.55, 0.61)},
    "صعب (D)":          {"Range": (0.70, 1.00), "Mean": (0.80, 0.88)},
    "يدوي":             {"Range": None,          "Mean": None},
}

# ── Mode 2: educational / curriculum assessments ─────────────────────────────
MODE2_EXAMS: dict[str, dict[str, int]] = {
    "القدرة المعرفية": {
        "المترادفات والمتضادات": 2, "الاستدلال اللفظي": 2,
        "الاستيعاب اللفظي": 2, "العمليات الحسابية": 3,
        "سلاسل الأرقام": 2, "تفسير البيانات": 2,
        "سلاسل الأشكال": 1, "تطابق الأشكال": 1,
        "فتح الصندوق": 2, "الاستنباطي": 3,
    },
    "نافس 4 أسئلة": {"1": 1, "2": 1, "3": 1, "4": 1},
    "نافس 5 أسئلة": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 1},
    "مخصص": {},
}

MODE2_STAGES: dict[str, dict] = {
    "علوم — الصف الثالث":       {"Range": (0.20, 0.80), "Mean": (0.50, 0.52)},
    "علوم — الصف السادس":       {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "علوم — الصف التاسع":       {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "الرياضيات — الصف الثالث":  {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "الرياضيات — الصف السادس":  {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "الرياضيات — الصف التاسع":  {"Range": (0.20, 0.80), "Mean": (0.64, 0.68)},
    "القراءة — الصف الثالث":    {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "القراءة — الصف السادس":    {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "القراءة — الصف التاسع":    {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "عام":                       {"Range": (0.20, 0.80), "Mean": (0.50, 0.52)},
    "يدوي":                      {"Range": None,          "Mean": None},
}

# ── Shared colours ────────────────────────────────────────────────────────────
HDR_BG  = "#1F3864"; HDR_FG = "#FFFFFF"
SUB_BG  = "#2E75B6"; ALT_BG = "#DCE6F1"
ACCENT  = "#ED7D31"
CHART_C = ["#2E75B6","#ED7D31","#70AD47","#C00000",
           "#7030A0","#00B0F0","#FF6600","#A9D18E"]

BIN_LABELS = [f"{i*0.1:.1f}–{(i+1)*0.1:.1f}" for i in range(10)]


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LOADER — MODE 1
# ══════════════════════════════════════════════════════════════════════════════

M1_COL_ALIASES = {
    "QuestionID": ["qoustionid","questionid","question_id","id","رقم","رقم السؤال"],
    "Category":   ["category","الفئة","التصنيف","الناتج","ناتج"],
    "D":          ["d","domain","المجال","مجال"],
    "Difficulty": ["difficulty","الصعوبة","صعوبة","b"],
    "تمييز":      ["تمييز","discrimination","a"],
    "مستوى التمييز": ["مستوى التمييز","مستوى","level","التخمين","التخميين","c"],
}


def _norm_cols(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
    lower = {c.strip().lower(): c for c in df.columns}
    rename = {}
    for canon, alts in aliases.items():
        if canon in df.columns:
            continue
        for a in alts:
            if a.strip().lower() in lower:
                rename[lower[a.strip().lower()]] = canon
                break
    return df.rename(columns=rename)


def load_bank_mode1(path: str | Path) -> pd.DataFrame:
    df = _read_file(path)
    df = _norm_cols(df, M1_COL_ALIASES)

    if "Category" not in df.columns:
        raise ValueError("العمود 'Category' (أو أي اسم مكافئ) غير موجود في الملف.")
    if "Difficulty" not in df.columns:
        raise ValueError("العمود 'Difficulty' غير موجود في الملف.")
    if "D" not in df.columns:
        df["D"] = "غير محدد"
    if "QuestionID" not in df.columns:
        df["QuestionID"] = [f"Q{i+1:05d}" for i in range(len(df))]

    df["Difficulty"] = pd.to_numeric(df["Difficulty"], errors="coerce").fillna(0.5).clip(0, 1)
    for col, default in [("تمييز", 1.0), ("مستوى التمييز", 0.25)]:
        if col not in df.columns:
            df[col] = default
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)

    df = df[~df["QuestionID"].duplicated(keep="first")].reset_index(drop=True)
    log.info("Mode 1 bank: %d questions loaded", len(df))
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LOADER — MODE 2
# ══════════════════════════════════════════════════════════════════════════════

M2_COL_ALIASES = {
    "QuestionID": ["qoustionid","questionid","question_id","id","رقم","رقم السؤال"],
    "الناتج":     ["الناتج","outcome","learning_outcome","ناتج","category"],
    "المؤشر":     ["المؤشر","indicator","مؤشر"],
    "المجال":     ["المجال","domain","d","مجال"],
    "difficulty": ["difficulty","الصعوبة","صعوبة","b"],
    "تمييز":      ["تمييز","discrimination","a"],
    "التخمين":    ["التخمين","التخميين","guessing","c","مستوى التمييز"],
    "subject":    ["subject","المادة"],
    "grade":      ["grade","الصف"],
    "language":   ["language","اللغة"],
}


def load_bank_mode2(path: str | Path) -> pd.DataFrame:
    df = _read_file(path)
    df.columns = [c.strip() for c in df.columns]      # strip whitespace
    df = _norm_cols(df, M2_COL_ALIASES)

    if "الناتج" not in df.columns:
        raise ValueError("العمود 'الناتج' (أو 'Category') غير موجود في الملف.")
    if "difficulty" not in df.columns:
        raise ValueError("العمود 'difficulty' غير موجود في الملف.")
    if "المجال" not in df.columns:
        df["المجال"] = "غير محدد"
    if "المؤشر" not in df.columns:
        df["المؤشر"] = ""
    if "QuestionID" not in df.columns:
        df["QuestionID"] = [f"Q{i+1:05d}" for i in range(len(df))]

    df["difficulty"] = pd.to_numeric(df["difficulty"], errors="coerce").fillna(0.5).clip(0, 1)
    for col, default in [("تمييز", 1.0), ("التخمين", 0.25)]:
        if col not in df.columns:
            df[col] = default
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)

    df["_pair_key"] = (
        df["الناتج"].astype(str) + "||" + df["المؤشر"].astype(str)
    )
    df = df[~df["QuestionID"].duplicated(keep="first")].reset_index(drop=True)
    log.info("Mode 2 bank: %d questions, %d pairs", len(df), df["_pair_key"].nunique())
    return df


def _read_file(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"الملف غير موجود: {path}")
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls", ".xlsm"):
        return pd.read_excel(path, engine="openpyxl")
    elif suffix == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig")
    else:
        raise ValueError(f"نوع الملف غير مدعوم: {suffix}")


# ══════════════════════════════════════════════════════════════════════════════
#  ASSEMBLY ENGINE (shared by both modes)
# ══════════════════════════════════════════════════════════════════════════════

def _largest_remainder(total: int, proportions: list[float]) -> list[int]:
    if not proportions or total == 0:
        return [0] * len(proportions)
    s = sum(proportions) or 1
    floats = [total * p / s for p in proportions]
    floors = [int(f) for f in floats]
    ranked = sorted(range(len(floats)), key=lambda i: -(floats[i] - floors[i]))
    deficit = total - sum(floors)
    for k in range(deficit):
        floors[ranked[k]] += 1
    return floors


@dataclass
class SlotSpec:
    """One selection slot: pick `count` items from `pool_filter` within `diff_range`."""
    label: str              # human label for logging
    pool_filter: dict       # {column: value} to filter bank rows
    count: int
    diff_range: tuple[float, float] = (0.0, 1.0)
    bins: list[tuple[float, float, int]] = field(default_factory=list)
    # bins: list of (low, high, count) — if provided, overrides count


@dataclass
class FormSpec:
    slots: list[SlotSpec]
    diff_range: tuple[float, float] = (0.0, 1.0)     # global filter
    diff_mean_range: Optional[tuple[float, float]] = None
    min_a: Optional[float] = None
    w_difficulty: float = 0.40
    w_bin:        float = 0.25
    w_usage:      float = 0.20
    w_random:     float = 0.15
    allow_reuse:  bool  = False
    max_reuse:    int   = 3
    rng_seed:     Optional[int] = None


class AssemblyEngine:
    def __init__(self, bank: pd.DataFrame, spec: FormSpec,
                 diff_col: str,
                 usage: Optional[dict[str, int]] = None):
        self.bank     = bank.copy().reset_index(drop=True)
        self.spec     = spec
        self.dcol     = diff_col          # "Difficulty" (mode1) or "difficulty" (mode2)
        self.rng      = np.random.default_rng(spec.rng_seed)
        self.usage: dict[str, int] = usage or {
            str(q): 0 for q in self.bank["QuestionID"]
        }
        self.warnings: list[str] = []

    # ── public ────────────────────────────────────────────────────────────────
    def assemble(self, form_idx: int) -> tuple[pd.DataFrame, list[str]]:
        selected_ids: set[str] = set()
        rows: list[pd.DataFrame] = []

        for slot in self.spec.slots:
            pool = self._pool(slot, selected_ids)

            if slot.bins:
                for (lo, hi, cnt) in slot.bins:
                    if cnt == 0:
                        continue
                    got, ids = self._pick(pool, cnt, lo, hi, slot, form_idx)
                    rows.extend(got); selected_ids.update(ids)
                    pool = pool[~pool["QuestionID"].isin(selected_ids)]
            else:
                got, ids = self._pick(pool, slot.count,
                                      slot.diff_range[0], slot.diff_range[1],
                                      slot, form_idx)
                rows.extend(got); selected_ids.update(ids)

        form = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        # Update usage
        for qid in selected_ids:
            self.usage[str(qid)] = self.usage.get(str(qid), 0) + 1

        return form, list(self.warnings)

    # ── pool ──────────────────────────────────────────────────────────────────
    def _pool(self, slot: SlotSpec, exclude: set[str]) -> pd.DataFrame:
        p = self.bank.copy()

        # Column filters (Category/D for mode1; الناتج/المؤشر for mode2)
        for col, val in slot.pool_filter.items():
            if col in p.columns and val:
                p = p[p[col].astype(str) == str(val)]

        # Exclude already chosen this form
        if exclude:
            p = p[~p["QuestionID"].isin(exclude)]

        # Over-use exclusion
        if not self.spec.allow_reuse:
            over = {q for q, c in self.usage.items() if c >= self.spec.max_reuse}
            p = p[~p["QuestionID"].isin(over)]

        # Discrimination filter
        if self.spec.min_a is not None and "تمييز" in p.columns:
            p = p[p["تمييز"] >= self.spec.min_a]

        # Global difficulty range
        lo, hi = self.spec.diff_range
        p = p[(p[self.dcol] >= lo) & (p[self.dcol] <= hi)]

        return p.reset_index(drop=True)

    # ── scoring ───────────────────────────────────────────────────────────────
    def _score(self, pool: pd.DataFrame, bin_lo: float, bin_hi: float) -> np.ndarray:
        sp   = self.spec
        diff = pool[self.dcol].to_numpy(float)
        qids = pool["QuestionID"].astype(str).tolist()

        # difficulty fit: closeness to midpoint of allowed range
        target = (sp.diff_range[0] + sp.diff_range[1]) / 2.0
        span   = max(sp.diff_range[1] - sp.diff_range[0], 1e-6)
        d_fit  = 1.0 - np.clip(np.abs(diff - target) / span, 0.0, 1.0)

        # bin fit
        bin_mid = (bin_lo + bin_hi) / 2.0
        hw      = max((bin_hi - bin_lo) / 2.0, 1e-6)
        b_fit   = np.clip(1.0 - (np.abs(diff - bin_mid) - hw) / hw, 0.0, 1.0)

        # usage penalty
        use_v = np.array([self.usage.get(q, 0) for q in qids], float)
        mu    = max(use_v.max(), 1.0)
        u_fit = 1.0 - (use_v / mu)

        noise = self.rng.random(len(pool))
        return (sp.w_difficulty * d_fit + sp.w_bin * b_fit
                + sp.w_usage * u_fit + sp.w_random * noise)

    # ── greedy pick ───────────────────────────────────────────────────────────
    def _pick(self, pool, needed, lo, hi, slot, form_idx):
        if pool.empty or needed == 0:
            return [], set()
        dcol = self.dcol

        # Stage 1: strict bin
        cands = pool[(pool[dcol] >= lo) & (pool[dcol] <= hi)].copy()
        # Stage 2: relax bin ±20 %
        if len(cands) < needed:
            hw    = (hi - lo) / 2.0
            cands = pool[(pool[dcol] >= lo - 0.2*hw) &
                         (pool[dcol] <= hi + 0.2*hw)].copy()
        # Stage 3: full pool (ignore bin)
        if len(cands) < needed:
            cands = pool.copy()
        # Stage 4: global bank (ignore slot filter)
        if cands.empty:
            used_set = set(self.usage.keys()) if not self.spec.allow_reuse else set()
            cands = self.bank[~self.bank["QuestionID"].isin(used_set)].copy()

        if cands.empty:
            self.warnings.append(f"نموذج {form_idx+1}: لا أسئلة متاحة — {slot.label}")
            return [], set()

        scores = self._score(cands, lo, hi)
        top    = cands.iloc[np.argsort(-scores)[:needed]]
        got    = len(top)
        if got < needed:
            self.warnings.append(
                f"نموذج {form_idx+1}: عجز {needed-got} سؤال — {slot.label}")
        return [top], set(top["QuestionID"].astype(str).tolist())


def run_assembly(bank: pd.DataFrame, spec: FormSpec,
                 diff_col: str, n_forms: int,
                 form_names: list[str]) -> list[tuple[pd.DataFrame, list[str]]]:
    """Assemble n_forms sharing usage counts so questions spread evenly."""
    usage = {str(q): 0 for q in bank["QuestionID"]}
    results = []
    for idx, name in enumerate(form_names):
        engine = AssemblyEngine(bank, spec, diff_col, usage)
        form, warns = engine.assemble(idx)
        usage = engine.usage
        mean_d = float(form[diff_col].mean()) if not form.empty else 0.0
        log.info("%s: %d items, mean_diff=%.3f", name, len(form), mean_d)
        results.append((form, warns))
    return results


# ══════════════════════════════════════════════════════════════════════════════
#  3PL PSYCHOMETRIC ANALYSIS (shared)
# ══════════════════════════════════════════════════════════════════════════════

def _p3pl(theta, a, b, c):
    theta = np.asarray(theta, float).reshape(-1, 1)
    a = np.asarray(a, float).reshape(1, -1)
    b = np.asarray(b, float).reshape(1, -1)
    c = np.asarray(c, float).reshape(1, -1)
    return np.clip(c + (1-c)/(1+np.exp(-a*(theta-b))), 1e-6, 1-1e-6)


def _item_info(theta, a, b, c):
    P  = _p3pl(theta, a, b, c)
    Q  = 1.0 - P
    a2 = np.asarray(a, float).reshape(1, -1)
    c2 = np.asarray(c, float).reshape(1, -1)
    return np.clip(a2**2 * ((P-c2)/(1-c2+1e-9))**2 * (Q/(P+1e-9)), 0, 100)


def _alpha(X):
    K = X.shape[0]
    if K < 2: return float("nan")
    tv = X.sum(axis=0).var(ddof=1)
    if tv == 0: return float("nan")
    return float(np.clip((K/(K-1))*(1-X.var(axis=1,ddof=1).sum()/tv), -1, 1))


def _item_corr(X):
    try:
        cm = np.corrcoef(X)
        return float(np.nanmean(cm[np.triu_indices_from(cm, k=1)]))
    except Exception:
        return float("nan")


@dataclass
class FormAnalysis:
    name: str
    n_items: int
    mean_diff: float; min_diff: float; max_diff: float
    avg_a: float; avg_b: float; avg_c: float
    alpha: float; item_corr: float
    theta_grid: np.ndarray      # (61,)
    icc: np.ndarray             # (61, K)
    tcc: np.ndarray             # (61,)
    tif: np.ndarray             # (61,)
    sem: np.ndarray             # (61,)
    score_hist: np.ndarray      # (n_students,)  — simulated total scores


def analyse_form(form_df: pd.DataFrame, name: str,
                 diff_col: str, a_col: str, c_col: str,
                 n_students: int = 3000,
                 rng_seed: Optional[int] = None) -> Optional[FormAnalysis]:
    if form_df.empty:
        return None
    df = form_df.copy()
    df["_b"] = pd.to_numeric(df.get(diff_col, 0.5), errors="coerce").fillna(0.5).clip(0, 1)
    df["_a"] = pd.to_numeric(df.get(a_col,    1.0), errors="coerce").fillna(1.0)
    df["_c"] = pd.to_numeric(df.get(c_col,    0.25),errors="coerce").fillna(0.25)
    df = df.dropna(subset=["_b"])
    if df.empty:
        return None
    b = df["_b"].to_numpy(float)
    a = df["_a"].to_numpy(float)
    c = df["_c"].to_numpy(float)
    K = len(df)

    theta_grid = np.linspace(-3, 3, 61)
    icc = _p3pl(theta_grid, a, b, c)          # (61, K)
    tcc = icc.sum(axis=1)
    tif = _item_info(theta_grid, a, b, c).sum(axis=1)
    sem = 1.0 / np.sqrt(np.clip(tif, 1e-8, None))

    rng     = np.random.default_rng(rng_seed)
    theta_s = rng.normal(0, 1, n_students)
    P_s     = _p3pl(theta_s, a, b, c)         # (n_students, K)
    X       = (rng.random((n_students, K)) < P_s).astype(np.float32).T  # (K, n_students)
    scores  = X.sum(axis=0)                    # (n_students,)
    al      = _alpha(X)
    ic      = _item_corr(X)

    return FormAnalysis(
        name=name, n_items=K,
        mean_diff=float(b.mean()), min_diff=float(b.min()), max_diff=float(b.max()),
        avg_a=float(a.mean()), avg_b=float(b.mean()), avg_c=float(c.mean()),
        alpha=al, item_corr=ic,
        theta_grid=theta_grid, icc=icc, tcc=tcc, tif=tif, sem=sem,
        score_hist=scores,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  OUTPUT — FORMS WORKBOOK
# ══════════════════════════════════════════════════════════════════════════════

def _fig_bytes(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


def _wb_fmts(wb):
    f = {}
    f["ttl"]  = wb.add_format({"bold":True,"font_size":13,"bg_color":HDR_BG,"font_color":HDR_FG,
                                "align":"center","valign":"vcenter","border":1})
    f["shdr"] = wb.add_format({"bold":True,"bg_color":SUB_BG,"font_color":HDR_FG,
                                "border":1,"align":"center","text_wrap":True})
    f["mlbl"] = wb.add_format({"bold":True,"bg_color":"#BDD7EE","border":1})
    f["mval"] = wb.add_format({"bg_color":"#FFF","border":1,"num_format":"0.0000"})
    f["mvi"]  = wb.add_format({"bg_color":"#FFF","border":1,"num_format":"0"})
    f["dat"]  = wb.add_format({"border":1,"align":"left","valign":"vcenter"})
    f["num"]  = wb.add_format({"border":1,"align":"center","num_format":"0.0000"})
    f["da2"]  = wb.add_format({"bg_color":ALT_BG,"border":1,"align":"left"})
    f["nu2"]  = wb.add_format({"bg_color":ALT_BG,"border":1,"align":"center","num_format":"0.0000"})
    f["red"]  = wb.add_format({"font_color":"red","bold":True,"border":1})
    f["warn"] = wb.add_format({"bg_color":"#FFEB9C","border":1,"font_color":"#9C5700"})
    return f


def write_forms_workbook(
    form_pairs: list[tuple[pd.DataFrame, list[str]]],
    form_names: list[str],
    analyses: list[Optional[FormAnalysis]],
    diff_col: str,
    a_col: str,
    c_col: str,
    usage_counts: dict[str, int],
    out_path: Path,
    mode: int,
):
    wb   = xlsxwriter.Workbook(str(out_path))
    fmts = _wb_fmts(wb)
    fa_map = {fa.name: fa for fa in analyses if fa}

    # decide which columns to export
    if mode == 1:
        export_cols = [
            ("QuestionID", 18, False), ("Category", 22, False),
            ("D", 18, False), ("Difficulty", 12, True),
            ("تمييز", 12, True), ("مستوى التمييز", 14, True),
        ]
    else:
        export_cols = [
            ("QuestionID", 18, False), ("الناتج", 22, False),
            ("المؤشر", 22, False), ("المجال", 18, False),
            ("difficulty", 12, True), ("تمييز", 12, True),
            ("التخمين", 12, True),
        ]

    for (form_df, warns), name, fa in zip(form_pairs, form_names, analyses):
        ws = wb.add_worksheet(name[:31])
        if mode == 2:
            ws.right_to_left()

        df = form_df.reset_index(drop=True)

        # ── Summary header block ────────────────────────────────────────────
        ws.merge_range("A1:D1", name, fmts["ttl"])
        ws.set_row(0, 22)
        metrics = [
            ("عدد الأسئلة",        len(df),                       False),
            ("متوسط الصعوبة",      df[diff_col].mean() if diff_col in df else "—", True),
            ("أدنى صعوبة",         df[diff_col].min()  if diff_col in df else "—", True),
            ("أقصى صعوبة",         df[diff_col].max()  if diff_col in df else "—", True),
            ("متوسط التمييز (a)",  df[a_col].mean()    if a_col in df else "—",    True),
            ("ألفا كرونباخ",       fa.alpha     if fa and not math.isnan(fa.alpha)     else "—", True),
            ("متوسط ارتباط البنود", fa.item_corr if fa and not math.isnan(fa.item_corr) else "—", True),
        ]
        ws.write(2, 0, "المقياس", fmts["shdr"])
        ws.write(2, 1, "القيمة",  fmts["shdr"])
        ws.set_column(0, 0, 28); ws.set_column(1, 1, 14)
        for ri, (lbl, val, is_f) in enumerate(metrics):
            r = 3 + ri
            ws.write(r, 0, lbl, fmts["mlbl"])
            if isinstance(val, str):
                ws.write(r, 1, val, fmts["dat"])
            elif is_f:
                try:    ws.write_number(r, 1, float(val), fmts["mval"])
                except: ws.write(r, 1, str(val), fmts["dat"])
            else:
                ws.write_number(r, 1, int(val), fmts["mvi"])

        # ── Warnings ────────────────────────────────────────────────────────
        warn_start = 11
        if warns:
            ws.write(warn_start, 0, "⚠ تحذيرات:", fmts["warn"])
            for wi, w in enumerate(warns[:5]):
                ws.write(warn_start+1+wi, 0, w, fmts["warn"])
            warn_start += len(warns[:5]) + 2

        # ── Item table ──────────────────────────────────────────────────────
        tbl = warn_start + 1
        for ci, (cn, cw, _) in enumerate(export_cols):
            ws.write(tbl, ci, cn, fmts["shdr"])
            ws.set_column(ci, ci, cw)

        for ri, (_, row) in enumerate(df.iterrows()):
            reused = usage_counts.get(str(row.get("QuestionID","")), 0) > 1
            for ci, (cn, _, is_n) in enumerate(export_cols):
                val = row.get(cn, "")
                if is_n:
                    fmt = fmts["nu2"] if ri%2 else fmts["num"]
                    try:    ws.write_number(tbl+1+ri, ci, float(val), fmt)
                    except: ws.write(tbl+1+ri, ci, str(val), fmt)
                else:
                    fmt = (fmts["red"] if reused
                           else (fmts["da2"] if ri%2 else fmts["dat"]))
                    ws.write(tbl+1+ri, ci, str(val) if val is not None else "", fmt)

        ws.freeze_panes(tbl+1, 0)
        ws.autofilter(tbl, 0, tbl+len(df), len(export_cols)-1)

        # ── Difficulty histogram (column chart) ─────────────────────────────
        _write_hist(wb, ws, df, diff_col, name[:31], col_offset=len(export_cols)+2, row_offset=tbl)

        # ── 3PL charts ──────────────────────────────────────────────────────
        if fa:
            img_row = tbl + len(df) + 4
            _embed_charts(wb, ws, fa, img_row)

    wb.close()
    log.info("Forms → %s", out_path)


def _write_hist(wb, ws, df, diff_col, sheet_name, col_offset, row_offset):
    if diff_col not in df.columns: return
    counts = [0]*10
    for d in df[diff_col].dropna():
        counts[min(int(float(d)*10), 9)] += 1
    fmts = _wb_fmts(wb)
    ws.write(row_offset, col_offset,   "نطاق الصعوبة", fmts["shdr"])
    ws.write(row_offset, col_offset+1, "العدد",        fmts["shdr"])
    for i, (lbl, cnt) in enumerate(zip(BIN_LABELS, counts)):
        ws.write(row_offset+1+i, col_offset,   lbl)
        ws.write(row_offset+1+i, col_offset+1, cnt)
    chart = wb.add_chart({"type":"column"})
    chart.add_series({
        "categories": [sheet_name, row_offset+1, col_offset,   row_offset+10, col_offset],
        "values":     [sheet_name, row_offset+1, col_offset+1, row_offset+10, col_offset+1],
        "data_labels": {"value": True},
    })
    chart.set_title({"name":"توزيع الصعوبة"})
    chart.set_x_axis({"name":"النطاق"}); chart.set_y_axis({"name":"العدد"})
    ws.insert_chart(row_offset+12, col_offset, chart)


def _embed_charts(wb, ws, fa: FormAnalysis, img_row: int):
    # Score histogram
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.hist(fa.score_hist, bins=15, color="#2E75B6", edgecolor="white", alpha=0.85)
    ax.set(xlabel="الدرجة الكلية", ylabel="عدد الطلاب",
           title="توزيع الدرجات المحاكاة")
    ax.grid(alpha=0.3, axis="y"); fig.tight_layout()
    ws.insert_image(img_row, 0, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)

    # Average ICC
    avg_icc = fa.icc.mean(axis=1)
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot(fa.theta_grid, avg_icc, color="#ED7D31", lw=2)
    ax.set(xlabel="القدرة (θ)", ylabel="P(صحيح)",
           title="منحنى الخصائص المتوسط (ICC)", xlim=(-3,3), ylim=(0,1.05))
    ax.grid(alpha=0.3); fig.tight_layout()
    ws.insert_image(img_row, 7, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)

    # SEM
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot(fa.theta_grid, fa.sem, color="#70AD47", lw=2)
    ax.set(xlabel="القدرة (θ)", ylabel="خطأ القياس",
           title="الخطأ المعياري للقياس (SEM)", xlim=(-3,3))
    ax.grid(alpha=0.3); fig.tight_layout()
    ws.insert_image(img_row+20, 0, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)


# ══════════════════════════════════════════════════════════════════════════════
#  OUTPUT — 3PL ANALYSIS WORKBOOK
# ══════════════════════════════════════════════════════════════════════════════

def write_analysis_workbook(analyses: list[FormAnalysis], out_path: Path):
    wb = OPWorkbook()
    wb.remove(wb.active)

    # ── Summary ──────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Summary")
    hdrs = ["Form","N Items","Avg a","Avg b","Avg c",
            "Mean Diff","Min Diff","Max Diff","Cronbach α","Item Corr."]
    ws.append(hdrs)
    for fa in analyses:
        ws.append([
            fa.name, fa.n_items,
            round(fa.avg_a,4), round(fa.avg_b,4), round(fa.avg_c,4),
            round(fa.mean_diff,4), round(fa.min_diff,4), round(fa.max_diff,4),
            round(fa.alpha,4) if not math.isnan(fa.alpha) else "N/A",
            round(fa.item_corr,4) if not math.isnan(fa.item_corr) else "N/A",
        ])
    avgs = {f: float(np.nanmean([getattr(fa,f) for fa in analyses]))
            for f in ("avg_a","avg_b","avg_c","mean_diff","min_diff","max_diff","alpha","item_corr")}
    ws.append(["Overall Average","",
               round(avgs["avg_a"],4),round(avgs["avg_b"],4),round(avgs["avg_c"],4),
               round(avgs["mean_diff"],4),round(avgs["min_diff"],4),round(avgs["max_diff"],4),
               round(avgs["alpha"],4),round(avgs["item_corr"],4)])

    # ── ICC per form ─────────────────────────────────────────────────────────
    for fa in analyses:
        ws_icc = wb.create_sheet(f"{fa.name}_ICC"[:31])
        K = fa.icc.shape[1]
        ws_icc.append(["Theta"] + [f"Item {j+1}" for j in range(K)])
        for i, th in enumerate(fa.theta_grid):
            ws_icc.append([round(th,3)] + [round(float(fa.icc[i,j]),4) for j in range(K)])
        n = len(fa.theta_grid)
        chart = LineChart(); chart.title = f"ICC — {fa.name}"
        chart.y_axis.title = "P(θ)"; chart.x_axis.title = "Theta"; chart.style = 10
        dr = Reference(ws_icc, min_col=2, min_row=1, max_col=1+K, max_row=1+n)
        cr = Reference(ws_icc, min_col=1, min_row=2, max_row=1+n)
        chart.add_data(dr, titles_from_data=True); chart.set_categories(cr)
        ws_icc.add_chart(chart, f"{get_column_letter(K+3)}2")

    # ── TCC comparison ───────────────────────────────────────────────────────
    ws_tcc = wb.create_sheet("TCC_Comparison")
    ws_tcc.append(["Theta"] + [fa.name for fa in analyses])
    for i, th in enumerate(analyses[0].theta_grid):
        ws_tcc.append([round(th,3)] + [round(float(fa.tcc[i]),4) for fa in analyses])
    n = len(analyses[0].theta_grid)
    chart = LineChart(); chart.title = "TCC — All Forms"
    chart.y_axis.title = "Expected Score"; chart.x_axis.title = "Theta"; chart.style = 10
    dr = Reference(ws_tcc, min_col=2, min_row=1, max_col=1+len(analyses), max_row=1+n)
    cr = Reference(ws_tcc, min_col=1, min_row=2, max_row=1+n)
    chart.add_data(dr, titles_from_data=True); chart.set_categories(cr)
    ws_tcc.add_chart(chart, f"{get_column_letter(len(analyses)+3)}2")

    # ── TIF comparison ───────────────────────────────────────────────────────
    ws_tif = wb.create_sheet("TIF_Comparison")
    ws_tif.append(["Theta"] + [fa.name for fa in analyses])
    for i, th in enumerate(analyses[0].theta_grid):
        ws_tif.append([round(th,3)] + [round(float(fa.tif[i]),4) for fa in analyses])
    chart = LineChart(); chart.title = "TIF — All Forms"
    chart.y_axis.title = "Information"; chart.x_axis.title = "Theta"; chart.style = 10
    dr = Reference(ws_tif, min_col=2, min_row=1, max_col=1+len(analyses), max_row=1+n)
    cr = Reference(ws_tif, min_col=1, min_row=2, max_row=1+n)
    chart.add_data(dr, titles_from_data=True); chart.set_categories(cr)
    ws_tif.add_chart(chart, f"{get_column_letter(len(analyses)+3)}2")

    wb.save(str(out_path))
    log.info("3PL analysis → %s", out_path)


# ══════════════════════════════════════════════════════════════════════════════
#  GUI — SHARED WIDGETS
# ══════════════════════════════════════════════════════════════════════════════

FT = ("Segoe UI", 14, "bold"); FH = ("Segoe UI", 11, "bold")
FB = ("Segoe UI", 10);          FM = ("Consolas", 9)
BG_D = "#1F3864"; BG_M = "#2E75B6"; BG_L = "#EBF3FB"
TXD = "#1A1A2E"; TXL = "#FFFFFF"


def _entry(p, w=14, **kw):
    return tk.Entry(p, width=w, bg="#FFF", fg="#000",
                    font=FB, relief="solid", bd=1, **kw)


def _lbl(p, t, bold=False, **kw):
    return tk.Label(p, text=t, bg=p.cget("bg"),
                    fg=TXD, font=(FH if bold else FB), **kw)


def _btn(p, t, cmd, bg=BG_M, **kw):
    return tk.Button(p, text=t, command=cmd, bg=bg, fg=TXL, font=FB,
                     activebackground="#C55A11", activeforeground=TXL,
                     relief="flat", padx=12, pady=6, cursor="hand2", **kw)


def _scroll_frame(parent):
    outer  = tk.Frame(parent, bg=BG_L)
    canvas = tk.Canvas(outer, bg=BG_L, highlightthickness=0)
    vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    inner  = tk.Frame(canvas, bg=BG_L)
    wid    = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>",
               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>",
                lambda e: canvas.itemconfig(wid, width=e.width))
    canvas.bind_all("<MouseWheel>",
                    lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
    outer._inner = inner
    return outer


def _treeview(parent, columns, widths):
    s = ttk.Style()
    s.configure("G.Treeview.Heading", background=BG_M, foreground=TXL, font=FH)
    s.configure("G.Treeview", font=FB, rowheight=22)
    s.map("G.Treeview", background=[("selected", ACCENT)])
    frm  = tk.Frame(parent, bg=BG_L)
    frm.pack(fill="both", expand=True)
    tree = ttk.Treeview(frm, columns=columns, show="headings", style="G.Treeview")
    vsb  = ttk.Scrollbar(frm, orient="vertical",   command=tree.yview)
    hsb  = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    for col, w in zip(columns, widths):
        tree.heading(col, text=col)
        tree.column(col, width=w, minwidth=40, anchor="center")
    vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
    tree.grid(row=0, column=0, sticky="nsew")
    frm.columnconfigure(0, weight=1); frm.rowconfigure(0, weight=1)
    return tree


def _log_widget(parent):
    frm = tk.Frame(parent, bg=BG_L); frm.pack(fill="both", expand=True)
    box = tk.Text(frm, bg="#0D1117", fg="#58D68D", font=FM,
                  state="disabled", relief="flat", height=18)
    sb  = ttk.Scrollbar(frm, command=box.yview)
    box.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y"); box.pack(fill="both", expand=True)
    return box


def _append_log(box, msg, col=""):
    box.config(state="normal")
    tag = col or "g"; box.tag_config(tag, foreground=col or "#58D68D")
    box.insert("end", msg+"\n", tag); box.see("end")
    box.config(state="disabled")


# ══════════════════════════════════════════════════════════════════════════════
#  GUI — LAUNCHER SCREEN
# ══════════════════════════════════════════════════════════════════════════════

class LauncherWindow(tk.Tk):
    """
    Full-screen launch screen — user picks which assembly mode to open.
    Styled like a modern dashboard tile selector.
    """
    def __init__(self):
        super().__init__()
        self.title("STGS — نظام إنشاء الاختبارات الذكي")
        self.configure(bg=BG_D)
        self.resizable(True, True)
        self.minsize(820, 560)
        self._build()

    def _build(self):
        # ── Header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG_D, pady=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="STGS", bg=BG_D, fg=TXL,
                 font=("Segoe UI", 32, "bold")).pack()
        tk.Label(hdr, text="نظام إنشاء الاختبارات الذكي",
                 bg=BG_D, fg="#A0BADB", font=("Segoe UI", 13)).pack()

        # ── Subtitle ────────────────────────────────────────────────────────
        sub = tk.Frame(self, bg=BG_D, pady=4)
        sub.pack(fill="x")
        tk.Label(sub,
                 text="اختر نوع الاختبار الذي تريد إنشاءه",
                 bg=BG_D, fg="#C9D8ED", font=("Segoe UI", 11)).pack()

        # ── Mode tiles ──────────────────────────────────────────────────────
        tiles = tk.Frame(self, bg=BG_D, pady=30)
        tiles.pack(expand=True)

        self._make_tile(
            tiles,
            icon="📋",
            title="الوضع الأول",
            subtitle="القدرات والتحصيل",
            desc=(
                "الاختبارات الوطنية المعيارية\n"
                "قدرات علمي · قدرات نظري\n"
                "التحصيلي · القدرة المعرفية\n"
                "قدرات الجامعيين\n\n"
                "عمود التصنيف: Category\n"
                "عمود المجال: D\n"
                "مستويات الصعوبة: E · M · D · Stage 1"
            ),
            color="#2980B9",
            command=self._open_mode1,
            col=0,
        )

        # Divider
        tk.Frame(tiles, bg="#3A5070", width=2).grid(
            row=0, column=1, sticky="ns", padx=20, pady=10)

        self._make_tile(
            tiles,
            icon="🏫",
            title="الوضع الثاني",
            subtitle="التقييم التربوي والمناهج",
            desc=(
                "اختبارات المناهج والصفوف الدراسية\n"
                "علوم · رياضيات · قراءة\n"
                "الصف الثالث · السادس · التاسع\n"
                "نافس · القدرة المعرفية\n\n"
                "عمود الناتج: الناتج\n"
                "عمود المؤشر: المؤشر\n"
                "فلترة: مادة · صف · لغة"
            ),
            color="#27AE60",
            command=self._open_mode2,
            col=2,
        )

        # ── Footer ──────────────────────────────────────────────────────────
        ftr = tk.Frame(self, bg="#16284A", pady=10)
        ftr.pack(fill="x", side="bottom")
        tk.Label(ftr,
                 text="كلا الوضعين يشملان: محرك تجميع ذكي · تحليل 3PL تلقائي · إخراج Excel احترافي",
                 bg="#16284A", fg="#7FA8CC", font=("Segoe UI", 9)).pack()

    def _make_tile(self, parent, icon, title, subtitle, desc, color, command, col):
        outer = tk.Frame(parent, bg=color, bd=0, cursor="hand2")
        outer.grid(row=0, column=col, padx=10, pady=10, sticky="nsew")
        outer.bind("<Button-1>", lambda e: command())

        inner = tk.Frame(outer, bg=color, padx=24, pady=20)
        inner.pack(fill="both", expand=True)

        tk.Label(inner, text=icon, bg=color, fg=TXL,
                 font=("Segoe UI", 36)).pack(pady=(0, 4))
        tk.Label(inner, text=title, bg=color, fg=TXL,
                 font=("Segoe UI", 16, "bold")).pack()
        tk.Label(inner, text=subtitle, bg=color, fg="#D6EAF8",
                 font=("Segoe UI", 11, "italic")).pack(pady=(0, 10))
        tk.Label(inner, text=desc, bg=color, fg="#EBF5FB",
                 font=("Segoe UI", 9), justify="center").pack()

        btn_frame = tk.Frame(inner, bg=color); btn_frame.pack(pady=(16, 0))
        btn = tk.Button(btn_frame, text="فتح →", bg="#FFFFFF",
                        fg=color, font=("Segoe UI", 10, "bold"),
                        relief="flat", padx=18, pady=6, cursor="hand2",
                        command=command)
        btn.pack()

        # Hover effect
        for w in [outer, inner] + inner.winfo_children():
            try:
                darker = self._darken(color)
                w.bind("<Enter>", lambda e, f=outer, d=darker: f.config(bg=d))
                w.bind("<Leave>", lambda e, f=outer, c=color: f.config(bg=c))
            except Exception:
                pass

        parent.columnconfigure(col, weight=1)
        parent.rowconfigure(0, weight=1)

    @staticmethod
    def _darken(hex_color: str) -> str:
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        factor = 0.82
        return f"#{int(r*factor):02x}{int(g*factor):02x}{int(b*factor):02x}"

    def _open_mode1(self):
        self.withdraw()
        win = Mode1Window(on_close=self.deiconify)
        win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), self.deiconify()))

    def _open_mode2(self):
        self.withdraw()
        win = Mode2Window(on_close=self.deiconify)
        win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), self.deiconify()))


# ══════════════════════════════════════════════════════════════════════════════
#  GUI — MODE 1 WINDOW
#  National / standardised exams (Category + D + Difficulty)
# ══════════════════════════════════════════════════════════════════════════════

class Mode1Window(tk.Toplevel):
    def __init__(self, on_close=None):
        super().__init__()
        self.title("STGS — الوضع الأول: القدرات والتحصيل")
        self.configure(bg=BG_D)
        self.resizable(True, True)
        self.minsize(960, 680)
        self._on_close = on_close
        self.bank_df: Optional[pd.DataFrame] = None
        self.source_path: Optional[Path] = None
        self.analyses: list[FormAnalysis] = []
        self._q: queue.Queue = queue.Queue()
        self._build()
        self.after(80, self._poll)

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build(self):
        # Title bar
        hf = tk.Frame(self, bg="#2980B9", pady=10); hf.pack(fill="x")
        tk.Label(hf, text="📋  الوضع الأول — القدرات والتحصيل",
                 bg="#2980B9", fg=TXL, font=FT).pack(side="left", padx=16)
        _btn(hf, "← العودة", self._close, bg=BG_D).pack(side="right", padx=12)

        # Notebook
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("M1.TNotebook", background="#2980B9", borderwidth=0)
        s.configure("M1.TNotebook.Tab", background=BG_M, foreground=TXL,
                    font=FH, padding=[12, 5])
        s.map("M1.TNotebook.Tab",
              background=[("selected","#2980B9")], foreground=[("selected",TXL)])
        nb = ttk.Notebook(self, style="M1.TNotebook")
        nb.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        t1 = _scroll_frame(nb); t2 = _scroll_frame(nb)
        t3 = tk.Frame(nb, bg=BG_L); t4 = _scroll_frame(nb)
        nb.add(t1, text="  البنك  "); nb.add(t2, text="  الإعداد  ")
        nb.add(t3, text="  التوليد  "); nb.add(t4, text="  النتائج  ")
        self._nb = nb

        self._build_bank_tab(t1._inner)
        self._build_config_tab(t2._inner)
        self._build_generate_tab(t3)
        self._build_results_tab(t4._inner)

        # Status bar
        sf = tk.Frame(self, bg="#1A2E4A", height=22); sf.pack(fill="x", side="bottom")
        self._sv = tk.StringVar(value="جاهز")
        tk.Label(sf, textvariable=self._sv, bg="#1A2E4A", fg="#A0BADB",
                 font=FM, anchor="w", padx=8).pack(fill="x")

    # ── Tab: Bank ─────────────────────────────────────────────────────────────
    def _build_bank_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "بنك الأسئلة", bold=True).pack(anchor="w", pady=(0, 6))

        uf = tk.Frame(p, bg=BG_L); uf.pack(fill="x", pady=4)
        self._bp = tk.StringVar(value="لم يتم اختيار ملف")
        tk.Label(uf, textvariable=self._bp, bg=BG_L, fg=BG_D,
                 font=FB, width=55, anchor="w").pack(side="left")
        _btn(uf, "تحميل Excel / CSV", self._browse_bank,
             bg="#2980B9").pack(side="left", padx=8)

        self._bstats = tk.StringVar(value="")
        tk.Label(p, textvariable=self._bstats, bg=BG_L, fg=BG_M,
                 font=FB, justify="left").pack(anchor="w", pady=4)

        _lbl(p, "الفئات المكتشفة في العمود Category", bold=True).pack(
            anchor="w", pady=(10, 4))
        self._cat_tree = _treeview(p,
            ["Category","المجال (D)","العدد","متوسط الصعوبة","أدنى","أقصى"],
            [160, 140, 65, 120, 70, 70])

    def _browse_bank(self):
        path = fd.askopenfilename(
            title="اختر بنك الأسئلة",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("الكل","*.*")])
        if not path: return
        self._sv.set("جاري التحميل…")
        threading.Thread(target=lambda: self._load_t(path), daemon=True).start()

    def _load_t(self, path):
        try:
            df = load_bank_mode1(path)
            self._q.put(("bank", df, path))
        except Exception as e:
            self._q.put(("err", str(e)))

    def _on_bank(self, df, path):
        self.bank_df     = df
        self.source_path = Path(path)
        self._bp.set(Path(path).name)
        n = len(df); pairs = df["Category"].nunique()
        self._bstats.set(
            f"الأسئلة: {n}   |   الفئات: {pairs}   |   "
            f"الصعوبة: [{df['Difficulty'].min():.2f} , {df['Difficulty'].max():.2f}]   |   "
            f"المتوسط: {df['Difficulty'].mean():.3f}")
        # Populate category tree
        for i in self._cat_tree.get_children(): self._cat_tree.delete(i)
        grp = df.groupby("Category")
        for cat, g in grp:
            dom = g["D"].mode().iloc[0] if "D" in g else ""
            self._cat_tree.insert("","end", values=[
                cat, dom, len(g),
                f"{g['Difficulty'].mean():.3f}",
                f"{g['Difficulty'].min():.3f}",
                f"{g['Difficulty'].max():.3f}"])
        self._populate_subdomain_entries()
        self._sv.set(f"تم التحميل: {n} سؤال")

    # ── Tab: Config ───────────────────────────────────────────────────────────
    def _build_config_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        p.columnconfigure(0, weight=1); p.columnconfigure(1, weight=1)

        # Exam template
        ef = tk.LabelFrame(p, text="قالب الاختبار",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        ef.grid(row=0, column=0, sticky="nsew", padx=(0,8), pady=4)

        r0 = tk.Frame(ef, bg=BG_L); r0.pack(fill="x")
        tk.Label(r0, text="الاختبار:", bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._exam_var = tk.StringVar(value=list(MODE1_EXAMS.keys())[0])
        cb = ttk.Combobox(r0, textvariable=self._exam_var,
                          values=list(MODE1_EXAMS.keys()),
                          state="readonly", width=24)
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda e: self._populate_subdomain_entries())
        _btn(r0, "تطبيق", self._populate_subdomain_entries, bg=BG_D).pack(
            side="left", padx=4)

        # Sub-header
        hdr = tk.Frame(ef, bg=BG_L); hdr.pack(fill="x", pady=(6,2))
        for i, (t,w) in enumerate([("الفئة (Category)", 22), ("عدد الأسئلة", 9)]):
            tk.Label(hdr, text=t, bg=BG_M, fg=TXL, font=FH,
                     width=w, relief="flat", padx=4).grid(row=0,column=i,padx=2)
        self._sub_frame = tk.Frame(ef, bg=BG_L); self._sub_frame.pack(fill="x")
        self._sub_rows: list[tuple[str,tk.Entry]] = []

        # Params
        pf = tk.LabelFrame(p, text="معاملات التجميع",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        pf.grid(row=0, column=1, sticky="nsew", padx=(8,0), pady=4)

        self._nforms = self._le(pf,"عدد النماذج:",0,"4")
        self._stage_var = tk.StringVar(value=list(MODE1_STAGES.keys())[0])
        self._stage_var.trace_add("write", self._on_stage)
        tk.Label(pf,text="مستوى الصعوبة:",bg=BG_L,fg=TXD,font=FB).grid(
            row=1,column=0,sticky="w",padx=(0,8),pady=2)
        ttk.Combobox(pf,textvariable=self._stage_var,
                     values=list(MODE1_STAGES.keys()),
                     state="readonly",width=22).grid(row=1,column=1,sticky="w",pady=2)
        self._rng_lbl = tk.Label(pf,text="",bg=BG_L,fg=TXD,font=FB)
        self._rng_ent = _entry(pf,w=20); self._rng_ent.config(state="disabled")
        self._mean_lbl = tk.Label(pf,text="",bg=BG_L,fg=TXD,font=FB)
        self._mean_ent = _entry(pf,w=20); self._mean_ent.config(state="disabled")
        self._rng_lbl.grid(row=2,column=0,sticky="w"); self._rng_ent.grid(row=2,column=1,sticky="w")
        self._mean_lbl.grid(row=3,column=0,sticky="w"); self._mean_ent.grid(row=3,column=1,sticky="w")

        self._stats_var = tk.BooleanVar(value=False)
        tk.Checkbutton(pf,text="فعّل فلتر التمييز (تمييز ≥ 0.85)",
                       variable=self._stats_var,bg=BG_L,fg=TXD,font=FB).grid(
            row=4,column=0,columnspan=2,sticky="w",pady=4)
        self._partial_var = tk.BooleanVar(value=True)
        tk.Checkbutton(pf,text="السماح بالتعبئة الجزئية",
                       variable=self._partial_var,bg=BG_L,fg=TXD,font=FB).grid(
            row=5,column=0,columnspan=2,sticky="w")
        self._reuse_var = tk.BooleanVar(value=False)
        tk.Checkbutton(pf,text="السماح بإعادة الاستخدام (حتى 3 مرات)",
                       variable=self._reuse_var,bg=BG_L,fg=TXD,font=FB).grid(
            row=6,column=0,columnspan=2,sticky="w",pady=4)
        self._sim_n  = self._le(pf,"طلاب المحاكاة (3PL):", 7, "3000")
        self._seed   = self._le(pf,"Seed (فارغ=عشوائي):",  8, "")

        # Scoring weights
        wf = tk.LabelFrame(p, text="أوزان التسجيل (المجموع ≈ 1.0)",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        wf.grid(row=1, column=0, columnspan=2, sticky="ew", pady=4)
        rf = tk.Frame(wf, bg=BG_L); rf.pack(fill="x")
        self._wd = self._le2(rf,"ملاءمة الصعوبة:",0,"0.40")
        self._wb = self._le2(rf,"ملاءمة الحاوية:", 1,"0.25")
        self._wu = self._le2(rf,"غرامة الاستخدام:",2,"0.20")
        self._wr = self._le2(rf,"العشوائية:",       3,"0.15")

        # Manual bins
        bf = tk.LabelFrame(p, text="التوزيع اليدوي للحاويات (0.0→1.0) — اختياري",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        bf.grid(row=2, column=0, columnspan=2, sticky="ew", pady=4)
        self._use_bins = tk.BooleanVar(value=False)
        tk.Checkbutton(bf, text="تفعيل الحاويات اليدوية",
                       variable=self._use_bins, bg=BG_L, fg=TXD, font=FB).pack(anchor="w")
        self._bins_container = tk.Frame(bf, bg=BG_L)
        self._bins_container.pack(fill="x")
        self._bins_map: dict[str, list] = {}

    def _populate_subdomain_entries(self, *_):
        for w in self._sub_frame.winfo_children(): w.destroy()
        self._sub_rows.clear()
        exam = self._exam_var.get()
        for cat, cnt in MODE1_EXAMS.get(exam, {}).items():
            frm = tk.Frame(self._sub_frame, bg=BG_L); frm.pack(fill="x", pady=1)
            tk.Label(frm, text=cat, bg=BG_L, fg=TXD, font=FB,
                     width=26, anchor="w").grid(row=0, column=0, padx=4)
            e = _entry(frm, w=8); e.insert(0, str(cnt))
            e.grid(row=0, column=1, padx=4)
            self._sub_rows.append((cat, e))
        # rebuild bins
        self._rebuild_bins()

    def _rebuild_bins(self):
        for w in self._bins_container.winfo_children(): w.destroy()
        self._bins_map.clear()
        for cat, _ in self._sub_rows:
            lf = tk.LabelFrame(self._bins_container, text=f"Category: {cat}",
                               bg=BG_L, fg=BG_D, font=FM, padx=6, pady=4)
            lf.pack(fill="x", pady=2)
            entries = []
            for rw in range(2):
                frm = tk.Frame(lf, bg=BG_L); frm.pack(fill="x")
                for bi in range(5):
                    idx = rw*5+bi
                    tk.Label(frm, text=BIN_LABELS[idx]+":",
                             bg=BG_L, fg=TXD, font=FM, width=9).grid(
                        row=0, column=bi*2, padx=2)
                    e = _entry(frm, w=4); e.grid(row=0, column=bi*2+1, padx=2)
                    entries.append(e)
            self._bins_map[cat] = entries

    def _on_stage(self, *_):
        is_manual = "يدوي" in self._stage_var.get()
        if is_manual:
            self._rng_lbl.config(text="النطاق (مثال: 0.2, 0.8):")
            self._rng_ent.config(state="normal")
            self._mean_lbl.config(text="المتوسط (مثال: 0.5, 0.6):")
            self._mean_ent.config(state="normal")
        else:
            self._rng_lbl.config(text=""); self._rng_ent.config(state="disabled")
            self._mean_lbl.config(text=""); self._mean_ent.config(state="disabled")

    # ── Tab: Generate ─────────────────────────────────────────────────────────
    def _build_generate_tab(self, p):
        p.configure(bg=BG_L)
        tk.Label(p, text="توليد النماذج وتحليل 3PL",
                 bg=BG_L, fg=TXD, font=FH).pack(anchor="w", padx=20, pady=(14,4))
        bf = tk.Frame(p, bg=BG_L, pady=6); bf.pack(padx=20, anchor="w")
        self._gbtn = _btn(bf, "▶  توليد وتحليل", self._run, bg="#2980B9")
        self._gbtn.config(font=("Segoe UI",11,"bold"), pady=8, padx=20)
        self._gbtn.pack(side="left", padx=(0,12))
        self._pv = tk.DoubleVar(value=0)
        ttk.Progressbar(bf, variable=self._pv, maximum=100,
                        length=320, mode="determinate").pack(side="left")
        lf = tk.LabelFrame(p, text="السجل", bg=BG_L, fg=BG_D,
                           font=FH, padx=8, pady=8)
        lf.pack(fill="both", expand=True, padx=20, pady=8)
        self._logbox = _log_widget(lf)

    # ── Tab: Results ──────────────────────────────────────────────────────────
    def _build_results_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "ملخص النتائج", bold=True).pack(anchor="w", pady=(0,8))
        self._rtree = _treeview(p,
            ["النموذج","الأسئلة","متوسط الصعوبة","أدنى","أقصى",
             "متوسط a","متوسط c","ألفا كرونباخ","ارتباط البنود"],
            [90,65,115,65,65,80,70,110,110])
        bf = tk.Frame(p, bg=BG_L); bf.pack(anchor="w", pady=8)
        _btn(bf, "فتح مجلد الإخراج", self._open_folder, bg="#2980B9").pack(side="left")
        _btn(bf, "تحديث", self._refresh_results, bg=BG_D).pack(side="left", padx=6)

    # ── Run logic ─────────────────────────────────────────────────────────────
    def _run(self):
        if self.bank_df is None:
            mb.showerror("خطأ","حمّل بنك الأسئلة أولاً."); return
        try:
            params = self._collect()
        except ValueError as e:
            mb.showerror("خطأ في الإعداد", str(e)); return
        self._gbtn.config(state="disabled"); self._pv.set(0)
        self._logbox.config(state="normal"); self._logbox.delete("1.0","end")
        self._logbox.config(state="disabled")
        threading.Thread(target=self._gen_t, args=(params,), daemon=True).start()

    def _collect(self) -> dict:
        def fi(e,n):
            try: return float(e.get().strip())
            except: raise ValueError(f"قيمة غير صالحة لـ '{n}'")
        def ii(e,n):
            try: return int(e.get().strip())
            except: raise ValueError(f"قيمة غير صالحة لـ '{n}'")

        n_forms = ii(self._nforms,"عدد النماذج")
        stage   = self._stage_var.get()
        crit    = MODE1_STAGES.get(stage, {})
        diff_range = crit.get("Range") or (0.0, 1.0)
        mean_range = crit.get("Mean")
        if "يدوي" in stage:
            try:
                parts = [float(x) for x in self._rng_ent.get().split(",")]
                if len(parts)==2: diff_range=(parts[0],parts[1])
            except: pass
            try:
                parts = [float(x) for x in self._mean_ent.get().split(",")]
                if len(parts)==2: mean_range=(parts[0],parts[1])
            except: pass

        # Build slots
        slots = []
        for cat, ent in self._sub_rows:
            try: cnt = int(ent.get())
            except: cnt = 0
            if cnt <= 0: continue
            bins = []
            if self._use_bins.get() and cat in self._bins_map:
                bin_counts = []
                for e in self._bins_map[cat]:
                    try: bin_counts.append(int(e.get()))
                    except: bin_counts.append(0)
                total_bc = sum(bin_counts)
                if total_bc > 0:
                    scaled = _largest_remainder(cnt, bin_counts)
                    for bi,(lbl,bc) in enumerate(zip(BIN_LABELS,scaled)):
                        if bc>0:
                            lo,hi = bi*0.1, (bi+1)*0.1
                            bins.append((lo,hi,bc))
            slots.append(SlotSpec(
                label=cat,
                pool_filter={"Category": cat},
                count=cnt,
                diff_range=diff_range,
                bins=bins,
            ))

        if not slots:
            raise ValueError("لا توجد فئات محددة. تأكد من اختيار قالب الاختبار.")

        spec = FormSpec(
            slots=slots,
            diff_range=diff_range,
            diff_mean_range=mean_range,
            min_a=0.85 if self._stats_var.get() else None,
            w_difficulty=fi(self._wd,"ملاءمة الصعوبة"),
            w_bin=fi(self._wb,"ملاءمة الحاوية"),
            w_usage=fi(self._wu,"غرامة الاستخدام"),
            w_random=fi(self._wr,"العشوائية"),
            allow_reuse=self._reuse_var.get(),
            max_reuse=3,
            rng_seed=int(self._seed.get().strip()) if self._seed.get().strip() else None,
        )
        seed_s = self._seed.get().strip()
        return {"n_forms":n_forms,"spec":spec,"diff_col":"Difficulty",
                "a_col":"تمييز","c_col":"مستوى التمييز",
                "n_students":ii(self._sim_n,"طلاب المحاكاة"),
                "seed":int(seed_s) if seed_s else None}

    def _gen_t(self, p):
        try:
            n = p["n_forms"]; spec = p["spec"]
            names = [f"Form_{i+1}" for i in range(n)]
            self._lg(f"بدء التجميع: {n} نماذج…")
            results = run_assembly(self.bank_df, spec, p["diff_col"], n, names)
            self._q.put(("prog",40))

            all_warns = []
            for (form,warns),name in zip(results,names):
                for w in warns: self._lg(f"  ⚠ {w}","yellow")
                all_warns.extend(warns)
                md = form[p["diff_col"]].mean() if not form.empty else 0
                self._lg(f"  {name}: {len(form)} أسئلة — متوسط الصعوبة={md:.3f}")

            self._lg("تشغيل تحليل 3PL…")
            forms_list = [r[0] for r in results]
            analyses = [
                analyse_form(fd,nm,p["diff_col"],p["a_col"],p["c_col"],
                             p["n_students"],p["seed"])
                for fd,nm in zip(forms_list,names)
            ]
            self._q.put(("prog",75))
            for fa in analyses:
                if fa:
                    self._lg(f"  {fa.name}: ألفا={fa.alpha:.3f}, "
                             f"ارتباط={fa.item_corr:.3f}")

            self._lg("كتابة ملفات الإخراج…")
            base   = self.source_path.parent
            prefix = self.source_path.stem
            fp = base/f"{prefix}_Forms.xlsx"
            ap = base/f"{prefix}_3PL_Analysis.xlsx"
            rp = base/f"{prefix}_Remaining_Questions.xlsx"

            usage = {}
            for form,_ in results:
                for qid in form.get("QuestionID",pd.Series()).astype(str):
                    usage[qid] = usage.get(qid,0)+1

            write_forms_workbook(results, names, analyses,
                                 p["diff_col"],p["a_col"],p["c_col"],
                                 usage, fp, mode=1)
            valid_fa = [fa for fa in analyses if fa]
            if valid_fa:
                write_analysis_workbook(valid_fa, ap)

            used_ids = set()
            for form,_ in results:
                used_ids.update(form["QuestionID"].astype(str).tolist())
            remaining = self.bank_df[~self.bank_df["QuestionID"].isin(used_ids)]
            remaining.drop(columns=[c for c in remaining.columns if c.startswith("_")],
                           errors="ignore").to_excel(str(rp),index=False,engine="openpyxl")

            self._q.put(("prog",100))
            self._lg(f"\n  📄 النماذج:          {fp}")
            self._lg(f"  📊 تحليل 3PL:        {ap}")
            self._lg(f"  📋 الأسئلة المتبقية:  {rp}")
            self._lg("\n✓ اكتمل التوليد بنجاح!", "lime")
            self._q.put(("done", valid_fa))

        except Exception as e:
            self._lg(f"\n✗ خطأ: {e}", "red")
            log.exception("Mode1 generation failed")
            self._q.put(("fail",))

    def _refresh_results(self):
        for i in self._rtree.get_children(): self._rtree.delete(i)
        for fa in self.analyses:
            self._rtree.insert("","end",values=[
                fa.name, fa.n_items,
                f"{fa.mean_diff:.3f}", f"{fa.min_diff:.3f}", f"{fa.max_diff:.3f}",
                f"{fa.avg_a:.3f}", f"{fa.avg_c:.3f}",
                f"{fa.alpha:.3f}" if not math.isnan(fa.alpha) else "—",
                f"{fa.item_corr:.3f}" if not math.isnan(fa.item_corr) else "—",
            ])

    def _open_folder(self):
        if self.source_path:
            d = str(self.source_path.parent)
            if sys.platform=="win32": subprocess.Popen(["explorer",d])
            elif sys.platform=="darwin": subprocess.Popen(["open",d])
            else: subprocess.Popen(["xdg-open",d])

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _le(self,p,lbl,row,default=""):
        tk.Label(p,text=lbl,bg=BG_L,fg=TXD,font=FB,anchor="w").grid(
            row=row,column=0,sticky="w",padx=(0,8),pady=2)
        e=_entry(p,w=16); e.insert(0,default)
        e.grid(row=row,column=1,sticky="w",pady=2)
        return e

    def _le2(self,p,lbl,col,default=""):
        tk.Label(p,text=lbl,bg=BG_L,fg=TXD,font=FB).grid(
            row=0,column=col*2,padx=(8,4),pady=2)
        e=_entry(p,w=8); e.insert(0,default)
        e.grid(row=0,column=col*2+1,padx=(0,8),pady=2)
        return e

    def _lg(self,msg,col=""):
        self._q.put(("log",msg,col))

    def _poll(self):
        try:
            while True:
                item=self._q.get_nowait(); k=item[0]
                if k=="bank":   self._on_bank(item[1],item[2])
                elif k=="log":  _append_log(self._logbox,item[1],item[2] if len(item)>2 else "")
                elif k=="prog": self._pv.set(item[1])
                elif k=="done":
                    self.analyses=item[1]; self._gbtn.config(state="normal")
                    self._pv.set(100); self._refresh_results(); self._nb.select(3)
                    self._sv.set("اكتمل التوليد.")
                    mb.showinfo("تمّ","تم توليد وتحليل جميع النماذج.\nتم حفظ الملفات.")
                elif k=="fail":
                    self._gbtn.config(state="normal"); self._sv.set("فشل — راجع السجل.")
                elif k=="err":
                    self._sv.set(f"خطأ: {item[1]}"); mb.showerror("خطأ",item[1])
        except queue.Empty: pass
        self.after(80, self._poll)

    def _close(self):
        self.destroy()
        if self._on_close: self._on_close()


# ══════════════════════════════════════════════════════════════════════════════
#  GUI — MODE 2 WINDOW
#  Educational / curriculum assessments (الناتج + المؤشر + filters)
# ══════════════════════════════════════════════════════════════════════════════

class Mode2Window(tk.Toplevel):
    def __init__(self, on_close=None):
        super().__init__()
        self.title("STGS — الوضع الثاني: التقييم التربوي والمناهج")
        self.configure(bg=BG_D)
        self.resizable(True, True)
        self.minsize(960, 680)
        self._on_close = on_close
        self.bank_df: Optional[pd.DataFrame] = None
        self._filtered_df: Optional[pd.DataFrame] = None
        self.source_path: Optional[Path] = None
        self.analyses: list[FormAnalysis] = []
        self._q: queue.Queue = queue.Queue()
        self._build()
        self.after(80, self._poll)

    def _build(self):
        # Title bar
        hf = tk.Frame(self, bg="#1E8449", pady=10); hf.pack(fill="x")
        tk.Label(hf, text="🏫  الوضع الثاني — التقييم التربوي والمناهج",
                 bg="#1E8449", fg=TXL, font=FT).pack(side="left", padx=16)
        _btn(hf, "← العودة", self._close, bg=BG_D).pack(side="right", padx=12)

        # Notebook
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("M2.TNotebook", background="#1E8449", borderwidth=0)
        s.configure("M2.TNotebook.Tab", background="#27AE60", foreground=TXL,
                    font=FH, padding=[12, 5])
        s.map("M2.TNotebook.Tab",
              background=[("selected","#1E8449")], foreground=[("selected",TXL)])
        nb = ttk.Notebook(self, style="M2.TNotebook")
        nb.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        t1=_scroll_frame(nb); t2=_scroll_frame(nb); t3=_scroll_frame(nb)
        t4=tk.Frame(nb,bg=BG_L); t5=_scroll_frame(nb)
        nb.add(t1,text="  البنك  "); nb.add(t2,text="  الإعداد  ")
        nb.add(t3,text="  الحاويات  "); nb.add(t4,text="  التوليد  ")
        nb.add(t5,text="  النتائج  ")
        self._nb = nb

        self._build_bank_tab(t1._inner)
        self._build_config_tab(t2._inner)
        self._build_bins_tab(t3._inner)
        self._build_generate_tab(t4)
        self._build_results_tab(t5._inner)

        sf = tk.Frame(self,bg="#1A3D28",height=22); sf.pack(fill="x",side="bottom")
        self._sv = tk.StringVar(value="جاهز")
        tk.Label(sf,textvariable=self._sv,bg="#1A3D28",fg="#A9DFBF",
                 font=FM,anchor="w",padx=8).pack(fill="x")

    # ── Tab: Bank ─────────────────────────────────────────────────────────────
    def _build_bank_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p,"بنك الأسئلة",bold=True).pack(anchor="w",pady=(0,6))

        uf=tk.Frame(p,bg=BG_L); uf.pack(fill="x",pady=4)
        self._bp=tk.StringVar(value="لم يتم اختيار ملف")
        tk.Label(uf,textvariable=self._bp,bg=BG_L,fg=BG_D,
                 font=FB,width=55,anchor="w").pack(side="left")
        _btn(uf,"تحميل Excel / CSV",self._browse_bank,bg="#27AE60").pack(side="left",padx=8)

        self._bstats=tk.StringVar(value="")
        tk.Label(p,textvariable=self._bstats,bg=BG_L,fg=BG_M,
                 font=FB,justify="left").pack(anchor="w",pady=4)

        # Filters — shown only when columns detected
        self._ffrm=tk.LabelFrame(p,text="فلترة البنك",bg=BG_L,fg=BG_D,font=FH,padx=8,pady=6)
        self._ffrm.pack(fill="x",pady=4); self._ffrm.pack_forget()

        self._fsub=tk.StringVar(value="All")
        self._fgrd=tk.StringVar(value="All")
        self._flng=tk.StringVar(value="All")
        for col,(i,(t,var,attr)) in enumerate(enumerate([
            ("المادة:",self._fsub,"_cb_sub"),
            ("الصف:",  self._fgrd,"_cb_grd"),
            ("اللغة:", self._flng,"_cb_lng"),
        ])):
            tk.Label(self._ffrm,text=t,bg=BG_L,fg=TXD,font=FB).grid(
                row=0,column=i*2,padx=(6,2))
            cb=ttk.Combobox(self._ffrm,textvariable=var,state="readonly",width=14)
            cb.grid(row=0,column=i*2+1,padx=(0,10))
            cb.bind("<<ComboboxSelected>>",lambda e:self._apply_filters())
            setattr(self,attr,cb)

        # Pair summary
        _lbl(p,"ملخص الأزواج (الناتج، المؤشر)",bold=True).pack(anchor="w",pady=(10,4))
        self._ptree=_treeview(p,
            ["الناتج","المؤشر","العدد","متوسط الصعوبة","أدنى","أقصى"],
            [180,180,65,115,70,70])

    def _browse_bank(self):
        path=fd.askopenfilename(title="اختر بنك الأسئلة",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("الكل","*.*")])
        if not path: return
        self._sv.set("جاري التحميل…")
        threading.Thread(target=lambda:self._load_t(path),daemon=True).start()

    def _load_t(self,path):
        try:
            df=load_bank_mode2(path); self._q.put(("bank",df,path))
        except Exception as e: self._q.put(("err",str(e)))

    def _on_bank(self,df,path):
        self.bank_df=df; self._filtered_df=df.copy()
        self.source_path=Path(path); self._bp.set(Path(path).name)
        has_f=all(c in df.columns for c in ("subject","grade","language"))
        if has_f:
            self._ffrm.pack(fill="x",pady=4)
            for col,cb in [("subject",self._cb_sub),("grade",self._cb_grd),
                           ("language",self._cb_lng)]:
                cb["values"]=["All"]+sorted(df[col].dropna().astype(str).unique())
        else:
            self._ffrm.pack_forget()
        self._update_stats(df); self._populate_pair_tree(df)
        self._populate_pair_requirements(df); self._populate_bins(df)
        self._sv.set(f"تم التحميل: {len(df)} سؤال")

    def _update_stats(self,df):
        self._bstats.set(
            f"الأسئلة: {len(df)}   |   الأزواج: {df['_pair_key'].nunique()}   |   "
            f"الصعوبة: [{df['difficulty'].min():.2f},{df['difficulty'].max():.2f}]   |   "
            f"المتوسط: {df['difficulty'].mean():.3f}")

    def _populate_pair_tree(self,df):
        for i in self._ptree.get_children(): self._ptree.delete(i)
        for _,row in df.groupby("_pair_key",sort=False).agg(
            الناتج=("الناتج","first"),المؤشر=("المؤشر","first"),
            count=("QuestionID","count"),mean_d=("difficulty","mean"),
            min_d=("difficulty","min"),max_d=("difficulty","max")
        ).reset_index(drop=True).iterrows():
            self._ptree.insert("","end",values=[
                row["الناتج"],row["المؤشر"],row["count"],
                f"{row['mean_d']:.3f}",f"{row['min_d']:.3f}",f"{row['max_d']:.3f}"])

    def _apply_filters(self):
        if self.bank_df is None: return
        df=self.bank_df.copy()
        for col,var in [("subject",self._fsub),("grade",self._fgrd),("language",self._flng)]:
            if col in df.columns and var.get()!="All":
                df=df[df[col].astype(str)==var.get()]
        self._filtered_df=df
        self._update_stats(df); self._populate_pair_tree(df)
        self._populate_pair_requirements(df)

    # ── Tab: Config ───────────────────────────────────────────────────────────
    def _build_config_tab(self,p):
        p.configure(bg=BG_L,padx=20,pady=16)
        p.columnconfigure(0,weight=1); p.columnconfigure(1,weight=1)

        # Exam template
        ef=tk.LabelFrame(p,text="قالب الاختبار",bg=BG_L,fg=BG_D,font=FH,padx=10,pady=8)
        ef.grid(row=0,column=0,sticky="nsew",padx=(0,8),pady=4)
        r0=tk.Frame(ef,bg=BG_L); r0.pack(fill="x")
        tk.Label(r0,text="الاختبار:",bg=BG_L,fg=TXD,font=FB).pack(side="left")
        self._exam_var=tk.StringVar(value=list(MODE2_EXAMS.keys())[0])
        cb=ttk.Combobox(r0,textvariable=self._exam_var,
                        values=list(MODE2_EXAMS.keys()),state="readonly",width=22)
        cb.pack(side="left",padx=6)
        cb.bind("<<ComboboxSelected>>",lambda e:self._populate_pair_requirements(
            self._filtered_df or self.bank_df))
        _btn(r0,"تطبيق",lambda:self._populate_pair_requirements(
            self._filtered_df or self.bank_df),bg=BG_D).pack(side="left",padx=4)

        # Pair requirements from bank
        _lbl(ef,"متطلبات (الناتج، المؤشر)",bold=True).pack(anchor="w",pady=(8,2))
        hdr=tk.Frame(ef,bg=BG_L); hdr.pack(fill="x",pady=(0,2))
        for i,(t,w) in enumerate([("الناتج",18),("المؤشر",18),("العدد",7),("المتاح",7)]):
            tk.Label(hdr,text=t,bg=BG_M,fg=TXL,font=FH,width=w,
                     relief="flat",padx=4).grid(row=0,column=i,padx=2)
        self._pair_frm=tk.Frame(ef,bg=BG_L); self._pair_frm.pack(fill="x")
        self._pair_rows:list[dict]=[]
        self._pair_total=tk.Label(ef,text="",bg=BG_L,fg=TXD,font=FB)
        self._pair_total.pack(anchor="w",pady=4)

        # Parameters
        pf=tk.LabelFrame(p,text="معاملات التجميع",bg=BG_L,fg=BG_D,font=FH,padx=10,pady=8)
        pf.grid(row=0,column=1,sticky="nsew",padx=(8,0),pady=4)
        self._nforms=self._le(pf,"عدد النماذج:",0,"4")
        self._stage_var=tk.StringVar(value=list(MODE2_STAGES.keys())[0])
        self._stage_var.trace_add("write",self._on_stage)
        tk.Label(pf,text="المستوى:",bg=BG_L,fg=TXD,font=FB).grid(
            row=1,column=0,sticky="w",padx=(0,8),pady=2)
        ttk.Combobox(pf,textvariable=self._stage_var,
                     values=list(MODE2_STAGES.keys()),
                     state="readonly",width=26).grid(row=1,column=1,sticky="w",pady=2)
        self._rng_lbl=tk.Label(pf,text="",bg=BG_L,fg=TXD,font=FB)
        self._rng_ent=_entry(pf,w=20); self._rng_ent.config(state="disabled")
        self._mean_lbl=tk.Label(pf,text="",bg=BG_L,fg=TXD,font=FB)
        self._mean_ent=_entry(pf,w=20); self._mean_ent.config(state="disabled")
        self._rng_lbl.grid(row=2,column=0,sticky="w")
        self._rng_ent.grid(row=2,column=1,sticky="w")
        self._mean_lbl.grid(row=3,column=0,sticky="w")
        self._mean_ent.grid(row=3,column=1,sticky="w")
        self._stats_var=tk.BooleanVar(value=False)
        tk.Checkbutton(pf,text="فعّل فلتر التمييز (تمييز ≥ 0.5)",
                       variable=self._stats_var,bg=BG_L,fg=TXD,font=FB).grid(
            row=4,column=0,columnspan=2,sticky="w",pady=4)
        self._partial_var=tk.BooleanVar(value=True)
        tk.Checkbutton(pf,text="السماح بالتعبئة الجزئية",
                       variable=self._partial_var,bg=BG_L,fg=TXD,font=FB).grid(
            row=5,column=0,columnspan=2,sticky="w")
        self._reuse_var=tk.BooleanVar(value=False)
        tk.Checkbutton(pf,text="السماح بإعادة الاستخدام",
                       variable=self._reuse_var,bg=BG_L,fg=TXD,font=FB).grid(
            row=6,column=0,columnspan=2,sticky="w",pady=4)
        self._sim_n=self._le(pf,"طلاب المحاكاة:",7,"3000")
        self._seed=self._le(pf,"Seed:",8,"")

        # Weights
        wf=tk.LabelFrame(p,text="أوزان التسجيل",bg=BG_L,fg=BG_D,font=FH,padx=10,pady=8)
        wf.grid(row=1,column=0,columnspan=2,sticky="ew",pady=4)
        rf=tk.Frame(wf,bg=BG_L); rf.pack(fill="x")
        self._wd=self._le2(rf,"ملاءمة الصعوبة:",0,"0.40")
        self._wb=self._le2(rf,"ملاءمة الحاوية:", 1,"0.25")
        self._wu=self._le2(rf,"غرامة الاستخدام:",2,"0.20")
        self._wr=self._le2(rf,"العشوائية:",       3,"0.15")

    def _populate_pair_requirements(self,df):
        if df is None: return
        for w in self._pair_frm.winfo_children(): w.destroy()
        self._pair_rows.clear()
        pairs=df.groupby("_pair_key",sort=False).agg(
            الناتج=("الناتج","first"),المؤشر=("المؤشر","first"),
            count=("QuestionID","count")).reset_index(drop=True)
        n=len(pairs)
        if n==0: return
        try:
            total=sum(int(e.get()) for _,e in getattr(self,"_sub_rows",[])
                      if e.get().isdigit()) or 40
        except: total=40
        base=total//n; rem=total%n
        counts=[base+1 if i<rem else base for i in range(n)]
        for idx,(_,row) in enumerate(pairs.iterrows()):
            frm=tk.Frame(self._pair_frm,bg=BG_L); frm.pack(fill="x",pady=1)
            en=tk.Entry(frm,width=20,bg="#F0F0F0",fg=TXD,font=FB,state="readonly")
            em=tk.Entry(frm,width=20,bg="#F0F0F0",fg=TXD,font=FB,state="readonly")
            ec=_entry(frm,w=7)
            ea=tk.Entry(frm,width=7,bg="#F0F0F0",fg=TXD,font=FB,state="readonly")
            en.config(state="normal"); en.insert(0,str(row["الناتج"])); en.config(state="readonly")
            em.config(state="normal"); em.insert(0,str(row["المؤشر"])); em.config(state="readonly")
            ec.insert(0,str(counts[idx]))
            ea.config(state="normal"); ea.insert(0,str(int(row["count"]))); ea.config(state="readonly")
            for ci,w in enumerate([en,em,ec,ea]): w.grid(row=0,column=ci,padx=4)
            self._pair_rows.append({
                "pair_key":str(row["الناتج"])+"||"+str(row["المؤشر"]),
                "en":en,"em":em,"ec":ec,"ea":ea})
        self._upd_pair_total()

    def _upd_pair_total(self):
        total=sum(int(r["ec"].get()) for r in self._pair_rows
                  if r["ec"].get().strip().isdigit())
        col="green" if total>0 else "#C00000"
        self._pair_total.config(text=f"إجمالي الأسئلة المعيّنة: {total}",fg=col)

    def _on_stage(self,*_):
        is_m="يدوي" in self._stage_var.get()
        if is_m:
            self._rng_lbl.config(text="النطاق (مثال: 0.2, 0.8):")
            self._rng_ent.config(state="normal")
            self._mean_lbl.config(text="المتوسط (مثال: 0.5, 0.6):")
            self._mean_ent.config(state="normal")
        else:
            self._rng_lbl.config(text=""); self._rng_ent.config(state="disabled")
            self._mean_lbl.config(text=""); self._mean_ent.config(state="disabled")

    # ── Tab: Bins ─────────────────────────────────────────────────────────────
    def _build_bins_tab(self,p):
        p.configure(bg=BG_L,padx=20,pady=16)
        _lbl(p,"التوزيع اليدوي للحاويات",bold=True).pack(anchor="w",pady=(0,4))
        tk.Label(p,text="حدد عدد الأسئلة من كل نطاق صعوبة لكل زوج (الناتج، المؤشر).",
                 bg=BG_L,fg=BG_M,font=FB).pack(anchor="w")
        self._use_bins=tk.BooleanVar(value=False)
        tk.Checkbutton(p,text="تفعيل التوزيع اليدوي للحاويات",
                       variable=self._use_bins,bg=BG_L,fg=TXD,font=FH).pack(
            anchor="w",pady=6)
        self._bins_inner=tk.Frame(p,bg=BG_L); self._bins_inner.pack(fill="x")
        self._bins_map:dict[str,list]={}

    def _populate_bins(self,df):
        for w in self._bins_inner.winfo_children(): w.destroy()
        self._bins_map.clear()
        pairs=df.groupby("_pair_key",sort=False).agg(
            الناتج=("الناتج","first"),المؤشر=("المؤشر","first")
        ).reset_index(drop=True)
        for _,row in pairs.iterrows():
            pk=str(row["الناتج"])+"||"+str(row["المؤشر"])
            lf=tk.LabelFrame(self._bins_inner,
                              text=f"الناتج: {row['الناتج']}  /  المؤشر: {row['المؤشر']}",
                              bg=BG_L,fg=BG_D,font=FM,padx=6,pady=4)
            lf.pack(fill="x",pady=2)
            entries=[]
            for rw in range(2):
                frm=tk.Frame(lf,bg=BG_L); frm.pack(fill="x")
                for bi in range(5):
                    idx=rw*5+bi
                    tk.Label(frm,text=BIN_LABELS[idx]+":",bg=BG_L,fg=TXD,
                             font=FM,width=9).grid(row=0,column=bi*2,padx=2)
                    e=_entry(frm,w=4); e.grid(row=0,column=bi*2+1,padx=2)
                    entries.append(e)
            self._bins_map[pk]=entries

    # ── Tab: Generate ─────────────────────────────────────────────────────────
    def _build_generate_tab(self,p):
        p.configure(bg=BG_L)
        tk.Label(p,text="توليد النماذج وتحليل 3PL",bg=BG_L,fg=TXD,font=FH).pack(
            anchor="w",padx=20,pady=(14,4))
        bf=tk.Frame(p,bg=BG_L,pady=6); bf.pack(padx=20,anchor="w")
        self._gbtn=_btn(bf,"▶  توليد وتحليل",self._run,bg="#27AE60")
        self._gbtn.config(font=("Segoe UI",11,"bold"),pady=8,padx=20)
        self._gbtn.pack(side="left",padx=(0,12))
        self._pv=tk.DoubleVar(value=0)
        ttk.Progressbar(bf,variable=self._pv,maximum=100,
                        length=320,mode="determinate").pack(side="left")
        lf=tk.LabelFrame(p,text="السجل",bg=BG_L,fg=BG_D,font=FH,padx=8,pady=8)
        lf.pack(fill="both",expand=True,padx=20,pady=8)
        self._logbox=_log_widget(lf)

    # ── Tab: Results ──────────────────────────────────────────────────────────
    def _build_results_tab(self,p):
        p.configure(bg=BG_L,padx=20,pady=16)
        _lbl(p,"ملخص النتائج",bold=True).pack(anchor="w",pady=(0,8))
        self._rtree=_treeview(p,
            ["النموذج","الأسئلة","متوسط الصعوبة","أدنى","أقصى",
             "متوسط a","متوسط c","ألفا كرونباخ","ارتباط البنود"],
            [90,65,115,65,65,80,70,110,110])
        bf=tk.Frame(p,bg=BG_L); bf.pack(anchor="w",pady=8)
        _btn(bf,"فتح مجلد الإخراج",self._open_folder,bg="#27AE60").pack(side="left")
        _btn(bf,"تحديث",self._refresh_results,bg=BG_D).pack(side="left",padx=6)

    # ── Run logic ─────────────────────────────────────────────────────────────
    def _run(self):
        df=self._filtered_df or self.bank_df
        if df is None: mb.showerror("خطأ","حمّل بنك الأسئلة أولاً."); return
        try: params=self._collect(df)
        except ValueError as e: mb.showerror("خطأ في الإعداد",str(e)); return
        self._gbtn.config(state="disabled"); self._pv.set(0)
        self._logbox.config(state="normal"); self._logbox.delete("1.0","end")
        self._logbox.config(state="disabled")
        threading.Thread(target=self._gen_t,args=(df,params),daemon=True).start()

    def _collect(self,df) -> dict:
        def fi(e,n):
            try: return float(e.get().strip())
            except: raise ValueError(f"قيمة غير صالحة لـ '{n}'")
        def ii(e,n):
            try: return int(e.get().strip())
            except: raise ValueError(f"قيمة غير صالحة لـ '{n}'")

        n_forms=ii(self._nforms,"عدد النماذج")
        stage=self._stage_var.get(); crit=MODE2_STAGES.get(stage,{})
        diff_range=crit.get("Range") or (0.0,1.0)
        mean_range=crit.get("Mean")
        if "يدوي" in stage:
            try:
                parts=[float(x) for x in self._rng_ent.get().split(",")]
                if len(parts)==2: diff_range=(parts[0],parts[1])
            except: pass
            try:
                parts=[float(x) for x in self._mean_ent.get().split(",")]
                if len(parts)==2: mean_range=(parts[0],parts[1])
            except: pass

        slots=[]
        for row in self._pair_rows:
            try: cnt=int(row["ec"].get())
            except: cnt=0
            if cnt<=0: continue
            pk=row["pair_key"]
            natej=row["en"].get(); moshir=row["em"].get()
            bins=[]
            if self._use_bins.get() and pk in self._bins_map:
                bcs=[]
                for e in self._bins_map[pk]:
                    try: bcs.append(int(e.get()))
                    except: bcs.append(0)
                if sum(bcs)>0:
                    scaled=_largest_remainder(cnt,bcs)
                    for bi,(lbl,bc) in enumerate(zip(BIN_LABELS,scaled)):
                        if bc>0:
                            lo,hi=bi*0.1,(bi+1)*0.1
                            bins.append((lo,hi,bc))
            pf={"الناتج":natej}
            if moshir and moshir not in ("","None","nan"):
                pf["المؤشر"]=moshir
            slots.append(SlotSpec(
                label=f"{natej} / {moshir}",
                pool_filter=pf,
                count=cnt,
                diff_range=diff_range,
                bins=bins,
            ))

        if not slots:
            raise ValueError("لا توجد أزواج محددة. حمّل البنك وتأكد من الإعداد.")

        spec=FormSpec(
            slots=slots, diff_range=diff_range, diff_mean_range=mean_range,
            min_a=0.5 if self._stats_var.get() else None,
            w_difficulty=fi(self._wd,"ملاءمة الصعوبة"),
            w_bin=fi(self._wb,"ملاءمة الحاوية"),
            w_usage=fi(self._wu,"غرامة الاستخدام"),
            w_random=fi(self._wr,"العشوائية"),
            allow_reuse=self._reuse_var.get(), max_reuse=3,
            rng_seed=int(self._seed.get().strip()) if self._seed.get().strip() else None,
        )
        seed_s=self._seed.get().strip()
        return {"n_forms":n_forms,"spec":spec,"diff_col":"difficulty",
                "a_col":"تمييز","c_col":"التخمين",
                "n_students":ii(self._sim_n,"طلاب المحاكاة"),
                "seed":int(seed_s) if seed_s else None}

    def _gen_t(self,df,p):
        try:
            n=p["n_forms"]; spec=p["spec"]
            names=[f"Form_{i+1}" for i in range(n)]
            self._lg(f"بدء التجميع: {n} نماذج…")
            results=run_assembly(df,spec,p["diff_col"],n,names)
            self._q.put(("prog",40))

            for (form,warns),name in zip(results,names):
                for w in warns: self._lg(f"  ⚠ {w}","yellow")
                md=form[p["diff_col"]].mean() if not form.empty else 0
                self._lg(f"  {name}: {len(form)} أسئلة — متوسط الصعوبة={md:.3f}")

            self._lg("تشغيل تحليل 3PL…")
            analyses=[
                analyse_form(r[0],nm,p["diff_col"],p["a_col"],p["c_col"],
                             p["n_students"],p["seed"])
                for r,nm in zip(results,names)]
            self._q.put(("prog",75))
            for fa in analyses:
                if fa: self._lg(f"  {fa.name}: ألفا={fa.alpha:.3f}, ارتباط={fa.item_corr:.3f}")

            self._lg("كتابة ملفات الإخراج…")
            base=self.source_path.parent; prefix=self.source_path.stem
            fp=base/f"{prefix}_Forms.xlsx"
            ap=base/f"{prefix}_3PL_Analysis.xlsx"
            rp=base/f"{prefix}_Remaining_Questions.xlsx"

            usage={}
            for form,_ in results:
                for qid in form.get("QuestionID",pd.Series()).astype(str):
                    usage[qid]=usage.get(qid,0)+1

            write_forms_workbook(results,names,analyses,
                                 p["diff_col"],p["a_col"],p["c_col"],
                                 usage,fp,mode=2)
            valid_fa=[fa for fa in analyses if fa]
            if valid_fa: write_analysis_workbook(valid_fa,ap)

            used_ids=set()
            for form,_ in results: used_ids.update(form["QuestionID"].astype(str))
            remaining=df[~df["QuestionID"].isin(used_ids)]
            remaining.drop(columns=[c for c in remaining.columns if c.startswith("_")],
                           errors="ignore").to_excel(str(rp),index=False,engine="openpyxl")

            self._q.put(("prog",100))
            self._lg(f"\n  📄 النماذج:          {fp}")
            self._lg(f"  📊 تحليل 3PL:        {ap}")
            self._lg(f"  📋 الأسئلة المتبقية:  {rp}")
            self._lg("\n✓ اكتمل التوليد بنجاح!","lime")
            self._q.put(("done",valid_fa))

        except Exception as e:
            self._lg(f"\n✗ خطأ: {e}","red")
            log.exception("Mode2 generation failed")
            self._q.put(("fail",))

    def _refresh_results(self):
        for i in self._rtree.get_children(): self._rtree.delete(i)
        for fa in self.analyses:
            self._rtree.insert("","end",values=[
                fa.name,fa.n_items,
                f"{fa.mean_diff:.3f}",f"{fa.min_diff:.3f}",f"{fa.max_diff:.3f}",
                f"{fa.avg_a:.3f}",f"{fa.avg_c:.3f}",
                f"{fa.alpha:.3f}" if not math.isnan(fa.alpha) else "—",
                f"{fa.item_corr:.3f}" if not math.isnan(fa.item_corr) else "—",
            ])

    def _open_folder(self):
        if self.source_path:
            d=str(self.source_path.parent)
            if sys.platform=="win32": subprocess.Popen(["explorer",d])
            elif sys.platform=="darwin": subprocess.Popen(["open",d])
            else: subprocess.Popen(["xdg-open",d])

    def _le(self,p,lbl,row,default=""):
        tk.Label(p,text=lbl,bg=BG_L,fg=TXD,font=FB,anchor="w").grid(
            row=row,column=0,sticky="w",padx=(0,8),pady=2)
        e=_entry(p,w=16); e.insert(0,default)
        e.grid(row=row,column=1,sticky="w",pady=2)
        return e

    def _le2(self,p,lbl,col,default=""):
        tk.Label(p,text=lbl,bg=BG_L,fg=TXD,font=FB).grid(
            row=0,column=col*2,padx=(8,4),pady=2)
        e=_entry(p,w=8); e.insert(0,default)
        e.grid(row=0,column=col*2+1,padx=(0,8),pady=2)
        return e

    def _lg(self,msg,col=""):  self._q.put(("log",msg,col))

    def _poll(self):
        try:
            while True:
                item=self._q.get_nowait(); k=item[0]
                if k=="bank":   self._on_bank(item[1],item[2])
                elif k=="log":  _append_log(self._logbox,item[1],item[2] if len(item)>2 else "")
                elif k=="prog": self._pv.set(item[1])
                elif k=="done":
                    self.analyses=item[1]; self._gbtn.config(state="normal")
                    self._pv.set(100); self._refresh_results(); self._nb.select(4)
                    self._sv.set("اكتمل التوليد.")
                    mb.showinfo("تمّ","تم توليد وتحليل جميع النماذج.\nتم حفظ الملفات.")
                elif k=="fail":
                    self._gbtn.config(state="normal"); self._sv.set("فشل — راجع السجل.")
                elif k=="err":
                    self._sv.set(f"خطأ: {item[1]}"); mb.showerror("خطأ",item[1])
        except queue.Empty: pass
        self.after(80, self._poll)

    def _close(self):
        self.destroy()
        if self._on_close: self._on_close()


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    LauncherWindow().mainloop()
