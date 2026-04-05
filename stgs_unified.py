"""
STGS — Smart Test Generation System (Unified)
===============================================
Single-file application merging the logic of all three original scripts.

Improvements over the originals
---------------------------------
* Scoring-based item selection replaces random.sample + 100-retry loop.
* (الناتج, المؤشر) pairs treated as independent — numerically identical
  المؤشر values under different الناتج entries are never confused.
* Largest-Remainder Method for exact integer bin quotas (zero rounding drift).
* 3PL analysis runs automatically after form generation — no second file picker.
* 3000-student simulation (was 100 in Script 1).
* All typo column aliases accepted (e.g. "التخميين" / "التخمين" / "c").
* Difficulty scale: 0-1 classical (p-value), as in the originals.
* No scipy dependency.

Column names accepted (case-insensitive, strip spaces)
-------------------------------------------------------
QuestionID : QoustionID, questionid, question_id, رقم, رقم السؤال
الناتج     : الناتج, outcome, learning_outcome, ناتج
المؤشر     : المؤشر, indicator, مؤشر
المجال     : المجال, domain, d, مجال
difficulty : difficulty, الصعوبة, صعوبة, b
تمييز      : تمييز, discrimination, a
التخمين    : التخمين, التخميين, guessing, c
subject    : subject, المادة
grade      : grade, الصف
language   : language, اللغة
"""

from __future__ import annotations

# ── stdlib ──────────────────────────────────────────────────────────────────
import io
import logging
import math
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

# ── third-party (no scipy) ──────────────────────────────────────────────────
import matplotlib
matplotlib.use("Agg")
import matplotlib.cm as cm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-8s  %(message)s")
log = logging.getLogger("stgs")


# ════════════════════════════════════════════════════════════════════════════
#  SECTION 1 — CONSTANTS & CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════

# --- Predefined exam structures (merged from both scripts) ------------------
EXAM_STRUCTURES: dict[str, dict[str, int]] = {
    "قدرات علمي": {
        "MAR": 5, "MAL": 1, "MAN": 2, "MGE": 3,
        "VAN": 3, "VCA": 2, "VSC": 3, "VRC": 5
    },
    "قدرات نظري": {
        "حساب": 3, "بياني": 1, "هندسة": 2,
        "التناظر اللفظي": 5, "الخطأ السياقي": 4,
        "إكمال الجمل": 3, "استيعاب المقروء": 6
    },
    "القدرة المعرفية": {
        "المترادفات والمتضادات": 2, "الاستدلال اللفظي": 2,
        "الاستيعاب اللفظي": 2, "العمليات الحسابية": 3,
        "سلاسل الأرقام": 2, "تفسير البيانات": 2,
        "سلاسل الأشكال": 1, "تطابق الأشكال": 1,
        "فتح الصندوق": 2, "الاستنباطي": 3
    },
    "التحصيلي": {
        "أحياء": 18, "كيمياء": 18, "فيزياء": 18, "الرياضيات": 18
    },
    "قدرات الجامعيين": {
        "حساب": 2, "جبر": 2, "بياني": 2, "هندسة": 2,
        "CT1": 2, "CT2": 3, "CT3": 5,
        "التناظر اللفظي": 2, "خطاء السياقي": 2,
        "إستيعاب المقروء": 2, "إكمال الجمل": 2
    },
    "نافس 4 أسئلة": {"1": 1, "2": 1, "3": 1, "4": 1},
    "نافس 5 أسئلة": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 1},
    "مخصص (يدوي)": {},
}

# --- Difficulty stages/criteria (0-1 classical scale) ----------------------
DIFFICULTY_STAGES: dict[str, dict] = {
    "Stage1 — عام":            {"Range": (0.20, 0.90), "Mean": (0.52, 0.58)},
    "سهل (E)":                  {"Range": (0.00, 0.45), "Mean": (0.27, 0.35)},
    "متوسط (M)":                {"Range": (0.40, 0.75), "Mean": (0.55, 0.61)},
    "صعب (D)":                  {"Range": (0.70, 1.00), "Mean": (0.80, 0.88)},
    "علوم صف الثالث":           {"Range": (0.20, 0.80), "Mean": (0.50, 0.52)},
    "علوم صف السادس":           {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "علوم الصف التاسع":         {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "الرياضيات الصف الثالث":    {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "الرياضيات الصف السادس":    {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "الرياضيات الصف التاسع":    {"Range": (0.20, 0.80), "Mean": (0.64, 0.68)},
    "القراءة الصف الثالث":      {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "القراءة الصف السادس":      {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "القراءة الصف التاسع":      {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "يدوي (Manually set)":      {"Range": None, "Mean": None},
}

# --- Column alias registry --------------------------------------------------
COL_ALIASES: dict[str, list[str]] = {
    "QuestionID": ["qoustionid", "questionid", "question_id", "id",
                   "رقم السؤال", "رقم"],
    "الناتج":     ["الناتج", "outcome", "learning_outcome", "ناتج", "category"],
    "المؤشر":     ["المؤشر", "indicator", "مؤشر"],
    "المجال":     ["المجال", "domain", "d", "مجال"],
    "difficulty": ["difficulty", "الصعوبة", "صعوبة", "b"],
    "تمييز":      ["تمييز", "discrimination", "a"],
    "التخمين":    ["التخمين", "التخميين", "guessing", "c", "تخمين"],
    "subject":    ["subject", "المادة"],
    "grade":      ["grade", "الصف"],
    "language":   ["language", "اللغة"],
}

# IRT defaults (when columns missing)
IRT_DEFAULTS = {"تمييز": 1.0, "difficulty": 0.50, "التخمين": 0.25}

# Scoring weights for assembly engine
DEFAULT_WEIGHTS = {
    "w_difficulty": 0.40,
    "w_bin":        0.25,
    "w_usage":      0.20,
    "w_random":     0.15,
}


# ════════════════════════════════════════════════════════════════════════════
#  SECTION 2 — DATA LOADER
# ════════════════════════════════════════════════════════════════════════════

class DataLoadError(Exception):
    pass


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map raw column names to canonical names via COL_ALIASES."""
    rename: dict[str, str] = {}
    lower_map = {c.strip().lower(): c for c in df.columns}
    for canonical, aliases in COL_ALIASES.items():
        if canonical in df.columns:
            continue
        for alias in aliases:
            key = alias.strip().lower()
            if key in lower_map:
                rename[lower_map[key]] = canonical
                break
    return df.rename(columns=rename)


def _validate_bank(df: pd.DataFrame) -> pd.DataFrame:
    issues: list[str] = []

    # Require at least الناتج
    if "الناتج" not in df.columns:
        raise DataLoadError(
            "العمود 'الناتج' (أو 'Category') غير موجود في الملف.\n"
            "تأكد من أن الملف يحتوي على عمود باسم 'الناتج' أو 'Category'."
        )

    # Add QuestionID if missing
    if "QuestionID" not in df.columns:
        df = df.copy()
        df["QuestionID"] = [f"Q{i+1:05d}" for i in range(len(df))]
        issues.append("عمود QuestionID غير موجود — تم إنشاؤه تلقائياً.")

    # IRT defaults
    for col, default in IRT_DEFAULTS.items():
        if col not in df.columns:
            df[col] = default
            issues.append(f"عمود '{col}' غير موجود — تم تعيينه إلى {default}.")
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)

    # Clamp difficulty to [0, 1]
    bad = ((df["difficulty"] < 0) | (df["difficulty"] > 1)).sum()
    if bad:
        df["difficulty"] = df["difficulty"].clip(0.0, 1.0)
        issues.append(f"{bad} قيمة صعوبة خارج [0,1] — تم تقليصها.")

    # Deduplicate IDs
    dupes = df["QuestionID"].duplicated(keep="first").sum()
    if dupes:
        df = df[~df["QuestionID"].duplicated(keep="first")].copy()
        issues.append(f"{dupes} سطر مكرر في QuestionID — تم حذفها.")

    # Pair key
    if "المؤشر" in df.columns:
        df["_pair_key"] = (
            df["الناتج"].astype(str) + "||" + df["المؤشر"].astype(str)
        )
    else:
        df["_pair_key"] = df["الناتج"].astype(str) + "||__none__"

    if issues:
        log.warning("Bank validation:\n  " + "\n  ".join(issues))
    return df.reset_index(drop=True)


def load_bank(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise DataLoadError(f"الملف غير موجود: {path}")
    try:
        if path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
            raw = pd.read_excel(path, engine="openpyxl")
        elif path.suffix.lower() == ".csv":
            raw = pd.read_csv(path, encoding="utf-8-sig")
        else:
            raise DataLoadError(f"نوع الملف غير مدعوم: {path.suffix}")
    except DataLoadError:
        raise
    except Exception as exc:
        raise DataLoadError(f"فشل قراءة الملف: {exc}") from exc

    if raw.empty:
        raise DataLoadError("الملف فارغ.")

    df = _normalise_columns(raw)
    df = _validate_bank(df)
    log.info("Loaded %d questions from '%s'", len(df), path.name)
    return df


def get_pair_summary(df: pd.DataFrame) -> pd.DataFrame:
    grp = df.groupby("_pair_key", sort=False)
    return grp.agg(
        الناتج=("الناتج", "first"),
        المؤشر=(
            "المؤشر", "first"
        ) if "المؤشر" in df.columns else ("الناتج", "first"),
        count=("QuestionID", "count"),
        mean_diff=("difficulty", "mean"),
        min_diff=("difficulty", "min"),
        max_diff=("difficulty", "max"),
    ).reset_index(drop=True)


# ════════════════════════════════════════════════════════════════════════════
#  SECTION 3 — ASSEMBLY ENGINE
# ════════════════════════════════════════════════════════════════════════════

def _largest_remainder(total: int, proportions: list[float]) -> list[int]:
    if not proportions or total == 0:
        return [0] * len(proportions)
    s = sum(proportions)
    floats = [total * p / s for p in proportions]
    floors = [int(f) for f in floats]
    ranked = sorted(range(len(floats)),
                    key=lambda i: -(floats[i] - floors[i]))
    deficit = total - sum(floors)
    for k in range(deficit):
        floors[ranked[k]] += 1
    return floors


@dataclass
class BinSpec:
    low: float
    high: float
    count: int


@dataclass
class AssemblyRequest:
    """One row in the assembly plan: pick `count` items matching `pair_key`."""
    pair_key: str
    الناتج: str
    المؤشر: str
    count: int
    bins: list[BinSpec] = field(default_factory=list)


@dataclass
class AssemblyConfig:
    requests: list[AssemblyRequest]
    # Difficulty criteria (0-1 scale)
    diff_range: tuple[float, float] = (0.0, 1.0)
    diff_mean_range: Optional[tuple[float, float]] = None
    # IRT filter (optional)
    min_a: Optional[float] = None      # minimum تمييز
    min_disc: Optional[float] = None   # minimum مستوى التمييز
    # Scoring weights
    w_difficulty: float = 0.40
    w_bin:        float = 0.25
    w_usage:      float = 0.20
    w_random:     float = 0.15
    # Behaviour
    allow_reuse: bool = False
    max_reuse: int = 3
    rng_seed: Optional[int] = None


@dataclass
class AssemblyResult:
    form: pd.DataFrame
    warnings: list[str]
    mean_difficulty: float


class AssemblyEngine:
    """
    Scoring-based greedy assembly engine.

    For each (الناتج, المؤشر) pair × bin combination, items are scored:
        score = w_difficulty * difficulty_fit
              + w_bin        * bin_fit
              + w_usage      * (1 - normalised_usage)
              + w_random     * noise
    Items are selected by descending score — no random retry loops.
    Shortage is handled by 4-stage adaptive relaxation instead of placeholders.
    """

    def __init__(self, bank: pd.DataFrame, config: AssemblyConfig,
                 usage_counts: Optional[dict[str, int]] = None):
        self.bank = bank.copy().reset_index(drop=True)
        self.cfg  = config
        self.rng  = np.random.default_rng(config.rng_seed)
        self.usage: dict[str, int] = usage_counts or {
            str(q): 0 for q in self.bank["QuestionID"]
        }
        self.warnings: list[str] = []

    # ------------------------------------------------------------------ #
    def assemble_one_form(self, form_idx: int) -> AssemblyResult:
        selected_ids: set[str] = set()
        rows: list[pd.DataFrame] = []

        for req in self.cfg.requests:
            if not req.bins:
                # No bin subdivision — pick from full pair pool
                pool = self._pool(req.pair_key, selected_ids)
                got, ids = self._select(pool, req.count,
                                        BinSpec(self.cfg.diff_range[0],
                                                self.cfg.diff_range[1],
                                                req.count),
                                        selected_ids, form_idx, req)
                rows.extend(got); selected_ids.update(ids)
            else:
                for bspec in req.bins:
                    pool = self._pool(req.pair_key, selected_ids)
                    got, ids = self._select(pool, bspec.count, bspec,
                                            selected_ids, form_idx, req)
                    rows.extend(got); selected_ids.update(ids)

        form = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()

        # Update global usage counts
        for qid in selected_ids:
            self.usage[str(qid)] = self.usage.get(str(qid), 0) + 1

        mean_diff = float(form["difficulty"].mean()) if not form.empty else 0.0
        return AssemblyResult(form=form,
                              warnings=list(self.warnings),
                              mean_difficulty=mean_diff)

    # ------------------------------------------------------------------ #
    def _pool(self, pair_key: str, exclude: set[str]) -> pd.DataFrame:
        cfg = self.cfg
        pool = self.bank.copy()

        # Filter by pair
        if pair_key and "__none__" not in pair_key:
            pool = pool[pool["_pair_key"] == pair_key]

        # Exclude already selected this form
        if exclude:
            pool = pool[~pool["QuestionID"].isin(exclude)]

        # Exclude over-used items
        if not cfg.allow_reuse:
            used = {q for q, c in self.usage.items() if c >= cfg.max_reuse}
            pool = pool[~pool["QuestionID"].isin(used)]

        # IRT filter
        if cfg.min_a is not None and "تمييز" in pool.columns:
            pool = pool[pool["تمييز"] >= cfg.min_a]

        # Difficulty range
        pool = pool[(pool["difficulty"] >= cfg.diff_range[0]) &
                    (pool["difficulty"] <= cfg.diff_range[1])]

        return pool.reset_index(drop=True)

    # ------------------------------------------------------------------ #
    def _score(self, pool: pd.DataFrame, bspec: BinSpec) -> np.ndarray:
        cfg = self.cfg
        diff = pool["difficulty"].to_numpy(float)
        qids = pool["QuestionID"].astype(str).tolist()

        # Difficulty fit — proximity to midpoint of allowed range
        target = (cfg.diff_range[0] + cfg.diff_range[1]) / 2.0
        span   = max(cfg.diff_range[1] - cfg.diff_range[0], 1e-6)
        d_fit  = 1.0 - np.clip(np.abs(diff - target) / span, 0.0, 1.0)

        # Bin fit — soft score: 1 inside bin, decays outside
        mid = (bspec.low + bspec.high) / 2.0
        hw  = max((bspec.high - bspec.low) / 2.0, 1e-6)
        b_fit = np.clip(1.0 - (np.abs(diff - mid) - hw) / hw, 0.0, 1.0)

        # Usage penalty
        use_v = np.array([self.usage.get(q, 0) for q in qids], float)
        mu    = max(use_v.max(), 1.0)
        u_fit = 1.0 - (use_v / mu)

        noise = self.rng.random(len(pool))

        return (cfg.w_difficulty * d_fit + cfg.w_bin * b_fit
                + cfg.w_usage * u_fit + cfg.w_random * noise)

    # ------------------------------------------------------------------ #
    def _select(self, pool, needed, bspec, exclude, form_idx, req):
        if pool.empty or needed == 0:
            return [], set()

        # Stage 1: strict bin
        cands = pool[(pool["difficulty"] >= bspec.low) &
                     (pool["difficulty"] <= bspec.high)].copy()

        # Stage 2: relax bin edges ±20%
        if len(cands) < needed:
            hw = (bspec.high - bspec.low) / 2.0
            cands = pool[(pool["difficulty"] >= bspec.low - 0.2*hw) &
                         (pool["difficulty"] <= bspec.high + 0.2*hw)].copy()

        # Stage 3: full pair pool (ignore bin)
        if len(cands) < needed:
            cands = pool.copy()

        # Stage 4: global fallback (ignore pair + bin)
        if cands.empty:
            cands = self.bank[~self.bank["QuestionID"].isin(exclude)].copy()

        if cands.empty:
            self.warnings.append(
                f"نموذج {form_idx+1}: لا توجد أسئلة متاحة لـ '{req.الناتج}'"
                f" مؤشر '{req.المؤشر}' بن [{bspec.low:.2f},{bspec.high:.2f}]"
            )
            return [], set()

        scores = self._score(cands, bspec)
        top    = cands.iloc[np.argsort(-scores)[:needed]]

        got = len(top)
        if got < needed:
            self.warnings.append(
                f"نموذج {form_idx+1}: عجز {needed-got} سؤال لـ '{req.الناتج}'"
                f" مؤشر '{req.المؤشر}' بن [{bspec.low:.2f},{bspec.high:.2f}]"
            )

        ids = set(top["QuestionID"].astype(str).tolist())
        return [top], ids


# ── Multi-form coordinator ──────────────────────────────────────────────────

def assemble_forms(
    bank: pd.DataFrame,
    config: AssemblyConfig,
    n_forms: int,
    form_names: Optional[list[str]] = None,
) -> list[AssemblyResult]:
    """
    Assemble `n_forms` forms sequentially, sharing usage_counts so that
    questions are spread across forms rather than repeated.
    """
    if form_names is None:
        form_names = [f"Form_{i+1}" for i in range(n_forms)]

    usage: dict[str, int] = {str(q): 0 for q in bank["QuestionID"]}
    results: list[AssemblyResult] = []

    for idx, name in enumerate(form_names):
        engine = AssemblyEngine(bank, config, usage_counts=usage)
        res    = engine.assemble_one_form(idx)

        # Verify mean criterion (with up to 3 adaptive relaxation passes)
        if config.diff_mean_range and not res.form.empty:
            mn, mx = config.diff_mean_range
            if not (mn <= res.mean_difficulty <= mx):
                # Try relaxing constraint by 20% each step
                for relax in [0.2, 0.4, 0.6]:
                    engine2 = AssemblyEngine(bank, config, usage_counts=dict(usage))
                    res2 = engine2.assemble_one_form(idx)
                    target = (mn + mx) / 2.0
                    if abs(res2.mean_difficulty - target) < abs(res.mean_difficulty - target):
                        res = res2
                    if mn - relax*0.1 <= res.mean_difficulty <= mx + relax*0.1:
                        break
                if not (mn - 0.06 <= res.mean_difficulty <= mx + 0.06):
                    res.warnings.append(
                        f"{name}: لم يُحقَّق معيار المتوسط "
                        f"(المستهدف {mn:.2f}-{mx:.2f}، الناتج {res.mean_difficulty:.3f})"
                    )

        usage = engine.usage
        results.append(res)
        log.info(
            "%s: %d items, mean_diff=%.3f, warnings=%d",
            name, len(res.form), res.mean_difficulty, len(res.warnings)
        )

    return results


# ════════════════════════════════════════════════════════════════════════════
#  SECTION 4 — 3PL PSYCHOMETRIC ANALYSIS
# ════════════════════════════════════════════════════════════════════════════

def _p3pl(theta: np.ndarray, a: np.ndarray,
          b: np.ndarray, c: np.ndarray) -> np.ndarray:
    """Vectorised 3PL — returns shape (N_theta, N_items)."""
    theta = np.asarray(theta, float).reshape(-1, 1)
    a = np.asarray(a, float).reshape(1, -1)
    b = np.asarray(b, float).reshape(1, -1)
    c = np.asarray(c, float).reshape(1, -1)
    return np.clip(c + (1.0 - c) / (1.0 + np.exp(-a * (theta - b))),
                   1e-6, 1.0 - 1e-6)


def _item_info(theta: np.ndarray, a: np.ndarray,
               b: np.ndarray, c: np.ndarray) -> np.ndarray:
    P  = _p3pl(theta, a, b, c)
    Q  = 1.0 - P
    a2 = np.asarray(a, float).reshape(1, -1)
    c2 = np.asarray(c, float).reshape(1, -1)
    return np.clip(a2**2 * ((P - c2) / (1.0 - c2 + 1e-9))**2 * (Q / P), 0, 100)


def _cronbach_alpha(X: np.ndarray) -> float:
    """KR-20 (binary) / Cronbach alpha from (n_items × n_students) matrix."""
    K = X.shape[0]
    if K < 2:
        return float("nan")
    total_var = X.sum(axis=0).var(ddof=1)
    if total_var == 0:
        return float("nan")
    return float(np.clip(
        (K / (K - 1)) * (1.0 - X.var(axis=1, ddof=1).sum() / total_var),
        -1.0, 1.0
    ))


def _item_correlation(X: np.ndarray) -> float:
    """Mean upper-triangle inter-item correlation."""
    try:
        cm = np.corrcoef(X)
        ut = cm[np.triu_indices_from(cm, k=1)]
        return float(np.nanmean(ut))
    except Exception:
        return float("nan")


@dataclass
class FormAnalysis:
    name: str
    n_items: int
    avg_a: float
    avg_b: float
    avg_c: float
    mean_diff: float
    min_diff: float
    max_diff: float
    alpha: float
    item_corr: float
    icc_curves: np.ndarray     # (n_theta, n_items)
    tcc: np.ndarray            # (n_theta,)
    tif: np.ndarray            # (n_theta,)
    sem: np.ndarray            # (n_theta,)
    theta_grid: np.ndarray     # (n_theta,)


def analyse_form(
    form_df: pd.DataFrame,
    name: str = "Form",
    n_students: int = 3000,
    rng_seed: Optional[int] = None,
) -> Optional[FormAnalysis]:
    """Full 3PL analysis for a single form DataFrame."""
    if form_df.empty:
        return None

    # Pick IRT columns — accept original typo "التخميين" or canonical "التخمين"
    b_col = "difficulty" if "difficulty" in form_df.columns else None
    a_col = "تمييز"      if "تمييز" in form_df.columns else None
    c_col = None
    for cc in ("التخمين", "التخميين", "c"):
        if cc in form_df.columns:
            c_col = cc; break

    if b_col is None:
        log.warning("'%s': عمود difficulty غير موجود — تخطّي", name)
        return None

    df = form_df.copy()
    df["_b"] = pd.to_numeric(df[b_col], errors="coerce").fillna(0.5).clip(0, 1)
    df["_a"] = pd.to_numeric(df[a_col], errors="coerce").fillna(1.0) if a_col else 1.0
    df["_c"] = pd.to_numeric(df[c_col], errors="coerce").fillna(0.25) if c_col else 0.25
    df = df.dropna(subset=["_b"])
    if df.empty:
        return None

    a = df["_a"].to_numpy(float)
    b = df["_b"].to_numpy(float)
    c = df["_c"].to_numpy(float)
    K = len(df)

    # Theta grid for curves
    theta_grid = np.linspace(-3.0, 3.0, 61)

    # ICC
    icc = _p3pl(theta_grid, a, b, c)       # (61, K)

    # TCC / TIF / SEM
    tcc = icc.sum(axis=1)
    tif_m = _item_info(theta_grid, a, b, c)
    tif  = tif_m.sum(axis=1)
    sem  = 1.0 / np.sqrt(np.clip(tif, 1e-8, None))

    # Simulate responses for alpha / correlation
    rng = np.random.default_rng(rng_seed)
    theta_s = rng.normal(0.0, 1.0, n_students)
    P_s = _p3pl(theta_s, a, b, c)          # (n_students, K)
    X   = (rng.random((n_students, K)) < P_s).astype(np.float32).T  # (K, n_students)

    alpha = _cronbach_alpha(X)
    icorr = _item_correlation(X)

    return FormAnalysis(
        name=name, n_items=K,
        avg_a=float(a.mean()), avg_b=float(b.mean()), avg_c=float(c.mean()),
        mean_diff=float(b.mean()), min_diff=float(b.min()), max_diff=float(b.max()),
        alpha=alpha, item_corr=icorr,
        icc_curves=icc, tcc=tcc, tif=tif, sem=sem, theta_grid=theta_grid,
    )


def analyse_all(
    forms: list[pd.DataFrame],
    names: list[str],
    n_students: int = 3000,
) -> list[FormAnalysis]:
    out = []
    for df, name in zip(forms, names):
        fa = analyse_form(df, name, n_students)
        if fa:
            out.append(fa)
    return out


# ════════════════════════════════════════════════════════════════════════════
#  SECTION 5 — OUTPUT GENERATOR
# ════════════════════════════════════════════════════════════════════════════

# Colour palette
HDR_BG  = "#1F3864"; HDR_FG = "#FFFFFF"
SUB_BG  = "#2E75B6"; ALT_BG = "#DCE6F1"
CHART_C = ["#2E75B6","#ED7D31","#A9D18E","#C00000",
           "#7030A0","#00B0F0","#70AD47","#FF6600"]

FORM_EXPORT_COLS = [
    ("QuestionID", 18, False),
    ("الناتج",    22, False),
    ("المؤشر",    22, False),
    ("المجال",    18, False),
    ("difficulty", 12, True),
    ("تمييز",     12, True),
    ("التخمين",   12, True),
]


def _wb_formats(wb) -> dict:
    f = {}
    f["ttl"]  = wb.add_format({"bold":True,"font_size":13,"bg_color":HDR_BG,
                                "font_color":HDR_FG,"align":"center","valign":"vcenter"})
    f["hdr"]  = wb.add_format({"bold":True,"bg_color":SUB_BG,"font_color":HDR_FG,
                                "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    f["mlbl"] = wb.add_format({"bold":True,"bg_color":"#BDD7EE","border":1})
    f["mval"] = wb.add_format({"bg_color":"#FFFFFF","border":1,"num_format":"0.0000"})
    f["mvi"]  = wb.add_format({"bg_color":"#FFFFFF","border":1,"num_format":"0"})
    f["dat"]  = wb.add_format({"border":1,"align":"left","valign":"vcenter"})
    f["num"]  = wb.add_format({"border":1,"align":"center","num_format":"0.0000"})
    f["da2"]  = wb.add_format({"bg_color":ALT_BG,"border":1,"align":"left"})
    f["nu2"]  = wb.add_format({"bg_color":ALT_BG,"border":1,"align":"center","num_format":"0.0000"})
    f["red"]  = wb.add_format({"font_color":"red","border":1})
    return f


def _fig_bytes(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


def _difficulty_histogram(writer, ws, form_df: pd.DataFrame,
                           sheet_name: str, col_offset: int,
                           row_offset: int) -> int:
    """Write bin table + embedded bar chart. Returns next free row."""
    wb = writer.book
    bins = np.linspace(0, 1, 11)
    labels = [f"{bins[i]:.1f}-{bins[i+1]:.1f}" for i in range(10)]
    counts = defaultdict(int)
    for d in form_df["difficulty"].dropna():
        idx = min(int(d * 10), 9)
        counts[labels[idx]] += 1

    ws.write(row_offset, col_offset, "نطاق الصعوبة", writer.book.add_format(
        {"bold":True,"bg_color":SUB_BG,"font_color":HDR_FG,"border":1}))
    ws.write(row_offset, col_offset+1, "العدد", writer.book.add_format(
        {"bold":True,"bg_color":SUB_BG,"font_color":HDR_FG,"border":1}))

    for i, lbl in enumerate(labels):
        ws.write(row_offset+1+i, col_offset, lbl)
        ws.write(row_offset+1+i, col_offset+1, counts[lbl])

    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "categories": [sheet_name, row_offset+1, col_offset,
                        row_offset+10, col_offset],
        "values":     [sheet_name, row_offset+1, col_offset+1,
                        row_offset+10, col_offset+1],
        "data_labels": {"value": True},
    })
    chart.set_title({"name": "توزيع الصعوبة"})
    chart.set_x_axis({"name": "النطاق"})
    chart.set_y_axis({"name": "العدد"})
    ws.insert_chart(row_offset+12, col_offset, chart)
    return row_offset + 28


def write_forms_file(
    results: list[AssemblyResult],
    form_names: list[str],
    analyses: list[Optional[FormAnalysis]],
    output_path: Path,
    usage_counts: dict[str, int],
) -> None:
    """Write one sheet per form with summary header, item table, and charts."""
    wb = xlsxwriter.Workbook(str(output_path))
    fmts = _wb_formats(wb)
    ana_map = {fa.name: fa for fa in analyses if fa}

    for res, name in zip(results, form_names):
        ws = wb.add_worksheet(name[:31])
        ws.right_to_left()
        df = res.form.reset_index(drop=True)
        fa = ana_map.get(name)

        # ── Summary header ──────────────────────────────────────────────
        ws.merge_range("A1:D1", name, fmts["ttl"])
        metrics = [
            ("عدد الأسئلة",       len(df),                      False),
            ("متوسط الصعوبة",     df["difficulty"].mean(),      True),
            ("أدنى صعوبة",        df["difficulty"].min(),       True),
            ("أقصى صعوبة",        df["difficulty"].max(),       True),
            ("متوسط التمييز (a)", df["تمييز"].mean()
             if "تمييز" in df.columns else 1.0,                  True),
            ("كرونباخ ألفا",      fa.alpha if fa else "—",      True),
            ("متوسط ارتباط البند", fa.item_corr if fa else "—", True),
        ]
        ws.write(2, 0, "المقياس", fmts["hdr"])
        ws.write(2, 1, "القيمة",  fmts["hdr"])
        for ri, (lbl, val, is_f) in enumerate(metrics):
            ws.write(3+ri, 0, lbl, fmts["mlbl"])
            if isinstance(val, str):
                ws.write(3+ri, 1, val, fmts["dat"])
            elif is_f and val is not None:
                try:
                    ws.write_number(3+ri, 1, float(val), fmts["mval"])
                except Exception:
                    ws.write(3+ri, 1, str(val), fmts["dat"])
            else:
                ws.write_number(3+ri, 1, int(val) if val else 0, fmts["mvi"])
        ws.set_column(0, 0, 28); ws.set_column(1, 1, 14)

        # ── Item table ──────────────────────────────────────────────────
        tbl_start = 12
        cols = list(FORM_EXPORT_COLS)
        if "Item Corr." not in df.columns and fa and fa.icc_curves is not None:
            pass  # item-level correlations not available per-row from 3PL sim
        for ci, (cn, w, _) in enumerate(cols):
            ws.write(tbl_start, ci, cn, fmts["hdr"])
            ws.set_column(ci, ci, w)

        for ri, (_, row) in enumerate(df.iterrows()):
            reused = usage_counts.get(str(row.get("QuestionID", "")), 0) > 1
            for ci, (cn, _, is_n) in enumerate(cols):
                val = row.get(cn, "")
                if is_n:
                    fmt = fmts["nu2"] if ri%2 else fmts["num"]
                    try:
                        ws.write_number(tbl_start+1+ri, ci, float(val), fmt)
                    except Exception:
                        ws.write(tbl_start+1+ri, ci, str(val), fmt)
                else:
                    fmt = (fmts["red"] if reused
                           else (fmts["da2"] if ri%2 else fmts["dat"]))
                    ws.write(tbl_start+1+ri, ci, str(val) if val is not None else "", fmt)

        ws.freeze_panes(tbl_start+1, 0)
        ws.autofilter(tbl_start, 0, tbl_start+len(df), len(cols)-1)

        # ── Difficulty histogram ─────────────────────────────────────────
        if not df.empty and "difficulty" in df.columns:
            _difficulty_histogram(writer=type('W', (), {'book': wb})(),
                                  ws=ws, form_df=df,
                                  sheet_name=name[:31],
                                  col_offset=10, row_offset=tbl_start)

        # ── 3PL charts (score dist, ICC, SEM) ───────────────────────────
        if fa:
            _embed_3pl_charts(wb, ws, fa, tbl_start + len(df) + 5)

    wb.close()
    log.info("Forms file → %s", output_path)


def _embed_3pl_charts(wb, ws, fa: FormAnalysis, img_row: int):
    """Embed ICC and SEM charts into the form worksheet."""
    # ICC (average)
    avg_icc = fa.icc_curves.mean(axis=1)
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot(fa.theta_grid, avg_icc, color="#2E75B6", lw=2)
    ax.set(xlabel="Ability (θ)", ylabel="P(correct)",
           title="Average ICC", xlim=(-3, 3), ylim=(0, 1.05))
    ax.grid(alpha=0.3); fig.tight_layout()
    ws.insert_image(img_row, 0, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)

    # SEM
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot(fa.theta_grid, fa.sem, color="#ED7D31", lw=2)
    ax.set(xlabel="Ability (θ)", ylabel="SEM",
           title="Standard Error of Measurement", xlim=(-3, 3))
    ax.grid(alpha=0.3); fig.tight_layout()
    ws.insert_image(img_row, 7, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)


def write_analysis_file(
    analyses: list[FormAnalysis],
    output_path: Path,
) -> None:
    """
    Write the 3PL analysis workbook (mirrors Script 3's output):
      - Summary sheet
      - ICC sheet per form (all item curves)
      - TCC comparison
      - TIF comparison
    Uses openpyxl for native Excel charts, same as Script 3.
    """
    from openpyxl import Workbook as OpWb
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import get_column_letter

    wb = OpWb()
    wb.remove(wb.active)

    # ── Summary sheet ────────────────────────────────────────────────────
    ws_sum = wb.create_sheet("Summary")
    headers = ["Form", "N Items", "Avg a", "Avg b", "Avg c",
               "Mean Difficulty", "Min Diff", "Max Diff",
               "Cronbach Alpha", "Item Corr."]
    ws_sum.append(headers)
    for fa in analyses:
        ws_sum.append([
            fa.name, fa.n_items,
            round(fa.avg_a, 4), round(fa.avg_b, 4), round(fa.avg_c, 4),
            round(fa.mean_diff, 4), round(fa.min_diff, 4), round(fa.max_diff, 4),
            round(fa.alpha, 4) if not math.isnan(fa.alpha) else "N/A",
            round(fa.item_corr, 4) if not math.isnan(fa.item_corr) else "N/A",
        ])
    # Overall average row
    numeric_cols = ["avg_a","avg_b","avg_c","mean_diff","min_diff","max_diff","alpha","item_corr"]
    avgs = {col: float(np.nanmean([getattr(fa, col) for fa in analyses]))
            for col in numeric_cols}
    ws_sum.append([
        "Overall Average", "",
        round(avgs["avg_a"],4), round(avgs["avg_b"],4), round(avgs["avg_c"],4),
        round(avgs["mean_diff"],4), round(avgs["min_diff"],4), round(avgs["max_diff"],4),
        round(avgs["alpha"],4), round(avgs["item_corr"],4),
    ])

    # ── ICC sheet per form ───────────────────────────────────────────────
    for fa in analyses:
        ws = wb.create_sheet(title=f"{fa.name}_ICC"[:31])
        K = fa.icc_curves.shape[1]
        item_labels = [f"Item {j+1}" for j in range(K)]
        ws.append(["Theta"] + item_labels)
        for i, theta in enumerate(fa.theta_grid):
            ws.append([round(theta, 3)] + [round(float(fa.icc_curves[i, j]), 4)
                                            for j in range(K)])
        n_rows = len(fa.theta_grid)
        chart = LineChart()
        chart.title = f"ICC – {fa.name}"
        chart.y_axis.title = "P(θ)"; chart.x_axis.title = "Theta"
        chart.style = 10
        data_ref = Reference(ws, min_col=2, min_row=1,
                              max_col=1+K, max_row=1+n_rows)
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=1+n_rows)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, f"{get_column_letter(K+3)}2")

    # ── TCC comparison ───────────────────────────────────────────────────
    ws_tcc = wb.create_sheet("TCC_Comparison")
    ws_tcc.append(["Theta"] + [fa.name for fa in analyses])
    for i, theta in enumerate(analyses[0].theta_grid):
        ws_tcc.append([round(theta, 3)] + [round(float(fa.tcc[i]), 4)
                                            for fa in analyses])
    n_rows = len(analyses[0].theta_grid)
    chart = LineChart()
    chart.title = "Test Characteristic Curves — All Forms"
    chart.y_axis.title = "Expected Score"; chart.x_axis.title = "Theta"
    chart.style = 10
    data_ref = Reference(ws_tcc, min_col=2, min_row=1,
                         max_col=1+len(analyses), max_row=1+n_rows)
    cats_ref = Reference(ws_tcc, min_col=1, min_row=2, max_row=1+n_rows)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws_tcc.add_chart(chart, f"{get_column_letter(len(analyses)+3)}2")

    # ── TIF comparison ───────────────────────────────────────────────────
    ws_tif = wb.create_sheet("TIF_Comparison")
    ws_tif.append(["Theta"] + [fa.name for fa in analyses])
    for i, theta in enumerate(analyses[0].theta_grid):
        ws_tif.append([round(theta, 3)] + [round(float(fa.tif[i]), 4)
                                            for fa in analyses])
    chart = LineChart()
    chart.title = "Test Information Function — All Forms"
    chart.y_axis.title = "Information"; chart.x_axis.title = "Theta"
    chart.style = 10
    data_ref = Reference(ws_tif, min_col=2, min_row=1,
                         max_col=1+len(analyses), max_row=1+n_rows)
    cats_ref = Reference(ws_tif, min_col=1, min_row=2, max_row=1+n_rows)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws_tif.add_chart(chart, f"{get_column_letter(len(analyses)+3)}2")

    wb.save(str(output_path))
    log.info("Analysis file → %s", output_path)


def write_remaining_file(bank: pd.DataFrame,
                         used_ids: set[str],
                         output_path: Path) -> None:
    remaining = bank[~bank["QuestionID"].isin(used_ids)].copy()
    remaining = remaining.drop(columns=["_pair_key"], errors="ignore")
    remaining.to_excel(str(output_path), index=False, engine="openpyxl")
    log.info("Remaining questions (%d) → %s", len(remaining), output_path)


# ════════════════════════════════════════════════════════════════════════════
#  SECTION 6 — GUI
# ════════════════════════════════════════════════════════════════════════════

# Theme
BG_D  = "#1F3864"; BG_M  = "#2E75B6"; BG_L  = "#EBF3FB"
ACC   = "#ED7D31";  TXD   = "#1A1A2E"; TXL   = "#FFFFFF"
FT = ("Segoe UI", 14, "bold"); FH = ("Segoe UI", 11, "bold")
FB = ("Segoe UI", 10);          FM = ("Consolas", 9)


def _e(p, w=12, **kw):
    return tk.Entry(p, width=w, bg="#FFF", fg="#000",
                    font=FB, relief="solid", bd=1, **kw)


def _lbl(p, t, bold=False, **kw):
    return tk.Label(p, text=t, bg=p.cget("bg"),
                    fg=TXD, font=(FH if bold else FB), **kw)


def _btn(p, t, cmd, bg=BG_M, **kw):
    return tk.Button(p, text=t, command=cmd, bg=bg, fg=TXL, font=FB,
                     activebackground="#C55A11", activeforeground=TXL,
                     relief="flat", padx=10, pady=5, cursor="hand2", **kw)


class STGS(tk.Tk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("STGS — نظام إنشاء الاختبارات الذكي")
        self.configure(bg=BG_D)
        self.resizable(True, True)
        self.minsize(980, 700)

        # State
        self.bank_df: Optional[pd.DataFrame] = None
        self.source_path: Optional[Path]     = None
        self.analyses: list[FormAnalysis]    = []
        self._q: queue.Queue                 = queue.Queue()

        # Per-form bin rows: dict[form_name, list[(Entry,Entry,Entry)]]
        self._bin_rows: dict[str, list] = {}
        # Manual stage criteria overrides
        self._manual_range: Optional[tuple] = None
        self._manual_mean:  Optional[tuple] = None

        self._build_header()
        self._build_notebook()
        self._build_statusbar()
        self.after(80, self._poll)

    # ── Header ──────────────────────────────────────────────────────────
    def _build_header(self):
        f = tk.Frame(self, bg=BG_D, pady=8)
        f.pack(fill="x")
        tk.Label(f, text="STGS — نظام إنشاء الاختبارات الذكي",
                 bg=BG_D, fg=TXL, font=FT).pack()

    # ── Notebook ────────────────────────────────────────────────────────
    def _build_notebook(self):
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("TNotebook", background=BG_D, borderwidth=0)
        s.configure("TNotebook.Tab", background=BG_M, foreground=TXL,
                    font=FH, padding=[14, 6])
        s.map("TNotebook.Tab",
              background=[("selected", ACC)], foreground=[("selected", TXL)])
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        t1 = self._scroll_tab(); t2 = self._scroll_tab()
        t3 = self._scroll_tab(); t4 = tk.Frame(self.nb, bg=BG_L)
        t5 = self._scroll_tab()

        self.nb.add(t1, text="  1. البنك  ")
        self.nb.add(t2, text="  2. الإعداد  ")
        self.nb.add(t3, text="  3. التوزيع اليدوي  ")
        self.nb.add(t4, text="  4. التوليد  ")
        self.nb.add(t5, text="  5. النتائج  ")

        self._build_bank_tab(t1._inner)
        self._build_config_tab(t2._inner)
        self._build_bins_tab(t3._inner)
        self._build_generate_tab(t4)
        self._build_results_tab(t5._inner)

    def _scroll_tab(self):
        outer  = tk.Frame(self.nb, bg=BG_L)
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
                        lambda e: canvas.yview_scroll(
                            int(-1*(e.delta/120)), "units"))
        outer._inner = inner
        return outer

    # ── Status bar ──────────────────────────────────────────────────────
    def _build_statusbar(self):
        f = tk.Frame(self, bg=BG_D, height=24)
        f.pack(fill="x", side="bottom")
        self._sv = tk.StringVar(value="جاهز")
        tk.Label(f, textvariable=self._sv, bg=BG_D, fg=TXL,
                 font=FM, anchor="w", padx=8).pack(fill="x")

    def _status(self, m):
        self._sv.set(m)

    # ════════════════════ TAB 1 — BANK ════════════════════════════════
    def _build_bank_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "بنك الأسئلة", bold=True).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        # Upload row
        uf = tk.Frame(p, bg=BG_L); uf.grid(row=1, column=0, sticky="ew", pady=4)
        self._bp = tk.StringVar(value="لم يتم اختيار ملف")
        tk.Label(uf, textvariable=self._bp, bg=BG_L, fg=BG_D,
                 font=FB, width=55, anchor="w").pack(side="left", padx=(0, 8))
        _btn(uf, "تحميل ملف Excel / CSV", self._browse_bank).pack(side="left")

        # Stats label
        self._bstats = tk.StringVar(value="")
        tk.Label(p, textvariable=self._bstats, bg=BG_L, fg=BG_M,
                 font=FB, justify="left").grid(
            row=2, column=0, sticky="w", pady=4)

        # Filters (shown only when subject/grade/language columns exist)
        self._filter_frame = tk.LabelFrame(
            p, text="فلترة البنك", bg=BG_L, fg=BG_D, font=FH, padx=8, pady=6)
        self._filter_frame.grid(row=3, column=0, sticky="ew", pady=4)
        self._filter_frame.grid_remove()

        self._fsubject  = tk.StringVar(value="All")
        self._fgrade    = tk.StringVar(value="All")
        self._flanguage = tk.StringVar(value="All")
        for i, (lbl, var, attr) in enumerate([
            ("المادة:", self._fsubject,  "_cb_subject"),
            ("الصف:",   self._fgrade,    "_cb_grade"),
            ("اللغة:",  self._flanguage, "_cb_language"),
        ]):
            tk.Label(self._filter_frame, text=lbl, bg=BG_L,
                     fg=TXD, font=FB).grid(row=0, column=i*2, padx=(6,2))
            cb = ttk.Combobox(self._filter_frame, textvariable=var,
                              state="readonly", width=14)
            cb.grid(row=0, column=i*2+1, padx=(0, 10))
            cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())
            setattr(self, attr, cb)

        # Pair summary tree
        _lbl(p, "ملخص (الناتج، المؤشر)", bold=True).grid(
            row=4, column=0, sticky="w", pady=(12, 4))
        tf = tk.Frame(p, bg=BG_L)
        tf.grid(row=5, column=0, sticky="nsew")
        self._ptree = self._treeview(
            tf,
            ["الناتج", "المؤشر", "العدد", "متوسط الصعوبة", "أدنى", "أقصى"],
            [180, 180, 65, 110, 70, 70],
        )

    def _browse_bank(self):
        path = fd.askopenfilename(
            title="اختر ملف بنك الأسئلة",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("الكل", "*.*")])
        if not path:
            return
        self._status("جاري تحميل البنك…")
        threading.Thread(target=self._load_bank_t,
                         args=(path,), daemon=True).start()

    def _load_bank_t(self, path):
        try:
            df = load_bank(path)
            self._q.put(("bank_ok", df, path))
        except Exception as e:
            self._q.put(("err", str(e)))

    def _on_bank(self, df: pd.DataFrame, path: str):
        self.bank_df       = df
        self.source_path   = Path(path)
        self._filtered_df  = df.copy()
        self._bp.set(Path(path).name)

        # Show filters if columns exist
        has_filters = all(c in df.columns for c in ("subject","grade","language"))
        if has_filters:
            self._filter_frame.grid()
            for col, cb in [("subject",  self._cb_subject),
                             ("grade",    self._cb_grade),
                             ("language", self._cb_language)]:
                vals = ["All"] + sorted(df[col].dropna().astype(str).unique())
                cb["values"] = vals
        else:
            self._filter_frame.grid_remove()

        self._update_bank_stats(df)
        self._populate_pair_tree(df)
        self._populate_pair_requirements(df)
        self._populate_bins_tab(df)
        self._status(f"تم التحميل: {len(df)} سؤال")

    def _update_bank_stats(self, df: pd.DataFrame):
        self._bstats.set(
            f"إجمالي الأسئلة: {len(df)}   |   "
            f"الأزواج (الناتج، المؤشر): {df['_pair_key'].nunique()}   |   "
            f"نطاق الصعوبة: [{df['difficulty'].min():.2f} , "
            f"{df['difficulty'].max():.2f}]   |   "
            f"متوسط الصعوبة: {df['difficulty'].mean():.3f}"
        )

    def _populate_pair_tree(self, df: pd.DataFrame):
        for i in self._ptree.get_children():
            self._ptree.delete(i)
        summ = get_pair_summary(df)
        for _, row in summ.iterrows():
            self._ptree.insert("", "end", values=[
                row["الناتج"],
                row.get("المؤشر", ""),
                row["count"],
                f"{row['mean_diff']:.3f}",
                f"{row['min_diff']:.3f}",
                f"{row['max_diff']:.3f}",
            ])

    def _apply_filters(self):
        if self.bank_df is None:
            return
        df = self.bank_df.copy()
        for col, var in [("subject",  self._fsubject),
                         ("grade",    self._fgrade),
                         ("language", self._flanguage)]:
            if col in df.columns and var.get() != "All":
                df = df[df[col].astype(str) == var.get()]
        self._filtered_df = df
        self._update_bank_stats(df)
        self._populate_pair_tree(df)
        self._populate_pair_requirements(df)

    # ════════════════════ TAB 2 — CONFIG ══════════════════════════════
    def _build_config_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        p.columnconfigure(0, weight=1); p.columnconfigure(1, weight=1)

        _lbl(p, "إعدادات التجميع", bold=True).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        # ── Basic params ────────────────────────────────────────────────
        bf = tk.LabelFrame(p, text="المعاملات الأساسية",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        bf.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=4)

        self._nforms  = self._le(bf, "عدد النماذج:", 0, "4")
        self._nstage  = self._stage_row(bf, 1)
        self._enable_stats = tk.BooleanVar(value=False)
        tk.Checkbutton(bf, text="تفعيل فلتر التمييز (تمييز ≥ 0.5)",
                       variable=self._enable_stats,
                       bg=BG_L, fg=TXD, font=FB).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=4)
        self._allow_partial = tk.BooleanVar(value=True)
        tk.Checkbutton(bf, text="السماح بالتعبئة الجزئية عند نقص الأسئلة",
                       variable=self._allow_partial,
                       bg=BG_L, fg=TXD, font=FB).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=4)
        self._allow_reuse = tk.BooleanVar(value=False)
        tk.Checkbutton(bf, text="السماح بإعادة استخدام الأسئلة عبر النماذج",
                       variable=self._allow_reuse,
                       bg=BG_L, fg=TXD, font=FB).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=4)
        self._sim_n = self._le(bf, "عدد الطلاب (المحاكاة):", 5, "3000")
        self._seed  = self._le(bf, "Seed (اتركه فارغاً=عشوائي):", 6, "")

        # ── Scoring weights ─────────────────────────────────────────────
        wf = tk.LabelFrame(p, text="أوزان التسجيل (مجموعها ≈ 1.0)",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        wf.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=4)
        self._wd = self._le(wf, "ملاءمة الصعوبة:", 0, "0.40")
        self._wb = self._le(wf, "ملاءمة الحاوية:", 1, "0.25")
        self._wu = self._le(wf, "غرامة الاستخدام:", 2, "0.20")
        self._wr = self._le(wf, "العشوائية:",        3, "0.15")

        # ── Exam template ────────────────────────────────────────────────
        ef = tk.LabelFrame(p, text="قالب الاختبار والتوزيع",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        ef.grid(row=2, column=0, columnspan=2, sticky="ew", pady=6)
        self._build_exam_template_section(ef)

        # ── (الناتج، المؤشر) requirements ────────────────────────────────
        pf = tk.LabelFrame(p, text="متطلبات (الناتج، المؤشر) — مستخرجة من البنك",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        pf.grid(row=3, column=0, columnspan=2, sticky="ew", pady=6)
        self._build_pair_req_section(pf)

    def _stage_row(self, parent, row):
        tk.Label(parent, text="مستوى الصعوبة:", bg=parent.cget("bg"),
                 fg=TXD, font=FB).grid(row=row, column=0, sticky="w",
                                       padx=(0, 8), pady=2)
        var = tk.StringVar(value=list(DIFFICULTY_STAGES.keys())[0])
        var.trace_add("write", self._on_stage_change)
        cb = ttk.Combobox(parent, textvariable=var,
                          values=list(DIFFICULTY_STAGES.keys()),
                          state="readonly", width=28)
        cb.grid(row=row, column=1, sticky="w", pady=2)
        self._stage_var = var

        # Manual override entries (hidden unless "يدوي")
        self._manual_range_lbl = tk.Label(
            parent, text="نطاق يدوي (مثال: 0.2, 0.8):",
            bg=BG_L, fg=TXD, font=FB)
        self._manual_range_ent = _e(parent, w=18)
        self._manual_mean_lbl  = tk.Label(
            parent, text="متوسط يدوي (مثال: 0.5, 0.6):",
            bg=BG_L, fg=TXD, font=FB)
        self._manual_mean_ent  = _e(parent, w=18)
        return var

    def _on_stage_change(self, *_):
        is_manual = "يدوي" in self._stage_var.get()
        for w in (self._manual_range_lbl, self._manual_range_ent,
                  self._manual_mean_lbl,  self._manual_mean_ent):
            if is_manual:
                w.grid(sticky="w", padx=20)
            else:
                w.grid_remove()

    def _build_exam_template_section(self, parent):
        row0 = tk.Frame(parent, bg=BG_L); row0.pack(fill="x", pady=(0, 6))
        tk.Label(row0, text="قالب الاختبار:", bg=BG_L, fg=TXD,
                 font=FB).pack(side="left", padx=(0, 8))
        self._exam_var = tk.StringVar(value=list(EXAM_STRUCTURES.keys())[0])
        cb = ttk.Combobox(row0, textvariable=self._exam_var,
                          values=list(EXAM_STRUCTURES.keys()),
                          state="readonly", width=30)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._populate_template_entries())
        _btn(row0, "تطبيق القالب", self._populate_template_entries,
             bg=BG_D).pack(side="left", padx=8)

        # Header
        hdr = tk.Frame(parent, bg=BG_L); hdr.pack(fill="x")
        for i, (t, w) in enumerate([("الناتج / المجال", 22), ("العدد", 8)]):
            tk.Label(hdr, text=t, bg=BG_M, fg=TXL, font=FH,
                     width=w, relief="flat", padx=4).grid(
                row=0, column=i, padx=2)

        self._tmpl_frame = tk.Frame(parent, bg=BG_L)
        self._tmpl_frame.pack(fill="x")
        self._tmpl_rows: list[tuple[str, tk.Entry]] = []

    def _populate_template_entries(self, *_):
        for w in self._tmpl_frame.winfo_children():
            w.destroy()
        self._tmpl_rows.clear()
        exam = self._exam_var.get()
        for subdomain, cnt in EXAM_STRUCTURES.get(exam, {}).items():
            frm = tk.Frame(self._tmpl_frame, bg=BG_L)
            frm.pack(fill="x", pady=1)
            tk.Label(frm, text=subdomain, bg=BG_L, fg=TXD,
                     font=FB, width=26, anchor="w").grid(row=0, column=0, padx=4)
            e = _e(frm, w=8); e.insert(0, str(cnt))
            e.grid(row=0, column=1, padx=4)
            self._tmpl_rows.append((subdomain, e))

    def _build_pair_req_section(self, parent):
        tk.Label(parent,
                 text="حمّل البنك أولاً لاستخراج الأزواج تلقائياً.",
                 bg=BG_L, fg=BG_M, font=FB).pack(anchor="w")
        hdr = tk.Frame(parent, bg=BG_L); hdr.pack(fill="x", pady=(4, 2))
        for i, (t, w) in enumerate([("الناتج", 20), ("المؤشر", 20),
                                     ("العدد", 8), ("المتاح", 8)]):
            tk.Label(hdr, text=t, bg=BG_M, fg=TXL, font=FH,
                     width=w, relief="flat", padx=4).grid(
                row=0, column=i, padx=2)
        self._pair_req_frame = tk.Frame(parent, bg=BG_L)
        self._pair_req_frame.pack(fill="x")
        self._pair_req_rows: list[dict] = []
        self._pair_total_lbl = tk.Label(parent, text="",
                                        bg=BG_L, fg=TXD, font=FB)
        self._pair_total_lbl.pack(anchor="w", pady=4)

    def _populate_pair_requirements(self, df: pd.DataFrame):
        for w in self._pair_req_frame.winfo_children():
            w.destroy()
        self._pair_req_rows.clear()
        summ = get_pair_summary(df)
        n = len(summ)
        if n == 0:
            return
        try:
            total = sum(int(e.get()) for _, e in self._tmpl_rows) or 40
        except Exception:
            total = 40
        base = total // n; rem = total % n
        counts = [base+1 if i < rem else base for i in range(n)]
        for idx, (_, row) in enumerate(summ.iterrows()):
            frm = tk.Frame(self._pair_req_frame, bg=BG_L)
            frm.pack(fill="x", pady=1)
            en = tk.Entry(frm, width=22, bg="#F0F0F0", fg=TXD,
                          font=FB, state="readonly")
            em = tk.Entry(frm, width=22, bg="#F0F0F0", fg=TXD,
                          font=FB, state="readonly")
            ec = _e(frm, w=8)
            ea = tk.Entry(frm, width=8, bg="#F0F0F0", fg=TXD,
                          font=FB, state="readonly")
            en.config(state="normal")
            en.insert(0, str(row["الناتج"]))
            en.config(state="readonly")
            em.config(state="normal")
            em.insert(0, str(row.get("المؤشر", "")))
            em.config(state="readonly")
            ec.insert(0, str(counts[idx]))
            ea.config(state="normal")
            ea.insert(0, str(int(row["count"])))
            ea.config(state="readonly")
            for ci, w in enumerate([en, em, ec, ea]):
                w.grid(row=0, column=ci, padx=4)
            self._pair_req_rows.append({
                "pair_key": str(row["الناتج"]) + "||" + str(row.get("المؤشر","")),
                "en": en, "em": em, "ec": ec, "ea": ea,
            })
        self._update_pair_total()

    def _update_pair_total(self):
        total = sum(int(r["ec"].get()) for r in self._pair_req_rows
                    if r["ec"].get().strip().isdigit())
        try:
            target = sum(int(e.get()) for _, e in self._tmpl_rows) or 40
        except Exception:
            target = 40
        col = "green" if total == target else "#C00000"
        self._pair_total_lbl.config(
            text=f"المجموع المعيّن: {total}  (المستهدف: {target})", fg=col)

    # ════════════════════ TAB 3 — MANUAL BINS ═════════════════════════
    def _build_bins_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "التوزيع اليدوي للصعوبة (حاويات 0.1)", bold=True).grid(
            row=0, column=0, sticky="w", pady=(0, 6))
        tk.Label(p, text="حدد عدد الأسئلة من كل نطاق صعوبة لكل زوج (الناتج، المؤشر).",
                 bg=BG_L, fg=BG_M, font=FB).grid(row=1, column=0, sticky="w")
        self._bins_tab_inner = tk.Frame(p, bg=BG_L)
        self._bins_tab_inner.grid(row=2, column=0, sticky="nsew", pady=8)
        self._bins_rows_map: dict[str, list] = {}  # pair_key → list of Entry
        self._use_manual_bins = tk.BooleanVar(value=False)
        tk.Checkbutton(p, text="تفعيل التوزيع اليدوي للحاويات",
                       variable=self._use_manual_bins,
                       bg=BG_L, fg=TXD, font=FH).grid(
            row=3, column=0, sticky="w", pady=6)

    def _populate_bins_tab(self, df: pd.DataFrame):
        for w in self._bins_tab_inner.winfo_children():
            w.destroy()
        self._bins_rows_map.clear()
        summ = get_pair_summary(df)
        bin_labels = [f"{i*0.1:.1f}-{(i+1)*0.1:.1f}" for i in range(10)]

        for row_idx, (_, row) in enumerate(summ.iterrows()):
            pair_key = str(row["الناتج"]) + "||" + str(row.get("المؤشر", ""))
            lf = tk.LabelFrame(
                self._bins_tab_inner,
                text=f"الناتج: {row['الناتج']}   المؤشر: {row.get('المؤشر','')}",
                bg=BG_L, fg=BG_D, font=FH, padx=8, pady=6)
            lf.grid(row=row_idx, column=0, sticky="ew", pady=4)

            entries = []
            # 2 rows × 5 bins
            for rw in range(2):
                frm = tk.Frame(lf, bg=BG_L); frm.pack(fill="x", pady=2)
                for bi in range(5):
                    idx = rw*5 + bi
                    tk.Label(frm, text=bin_labels[idx]+":",
                             bg=BG_L, fg=TXD, font=FM,
                             width=9).grid(row=0, column=bi*2, padx=2)
                    e = _e(frm, w=5)
                    e.grid(row=0, column=bi*2+1, padx=2)
                    entries.append(e)
            self._bins_rows_map[pair_key] = entries

    # ════════════════════ TAB 4 — GENERATE ════════════════════════════
    def _build_generate_tab(self, p):
        p.configure(bg=BG_L)
        tk.Label(p, text="توليد النماذج وتشغيل التحليل النفسي-التربوي",
                 bg=BG_L, fg=TXD, font=FH).pack(anchor="w", padx=20, pady=(14, 4))
        bf = tk.Frame(p, bg=BG_L, pady=6); bf.pack(padx=20, anchor="w")
        self._gbtn = _btn(bf, "▶  توليد وتحليل", self._run, bg=ACC)
        self._gbtn.config(font=("Segoe UI", 11, "bold"), pady=8, padx=20)
        self._gbtn.pack(side="left", padx=(0, 12))
        self._pv = tk.DoubleVar(value=0)
        ttk.Progressbar(bf, variable=self._pv, maximum=100,
                        length=320, mode="determinate").pack(side="left")
        lf = tk.LabelFrame(p, text="السجل", bg=BG_L, fg=BG_D,
                           font=FH, padx=8, pady=8)
        lf.pack(fill="both", expand=True, padx=20, pady=8)
        self._log = tk.Text(lf, height=22, bg="#0D1117", fg="#58D68D",
                            font=FM, state="disabled", relief="flat")
        sb = ttk.Scrollbar(lf, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True)

    # ════════════════════ TAB 5 — RESULTS ═════════════════════════════
    def _build_results_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        p.columnconfigure(0, weight=1)
        _lbl(p, "ملخص النتائج", bold=True).grid(
            row=0, column=0, sticky="w", pady=(0, 8))
        self._rtree = self._treeview(
            p,
            ["النموذج", "الأسئلة", "متوسط الصعوبة", "أدنى", "أقصى",
             "متوسط a", "متوسط c", "ألفا كرونباخ", "متوسط الارتباط"],
            [100, 70, 120, 70, 70, 80, 70, 110, 120],
            row=1,
        )
        bf = tk.Frame(p, bg=BG_L); bf.grid(row=2, column=0, sticky="w", pady=8)
        _btn(bf, "فتح مجلد الإخراج", self._open_folder).pack(side="left")
        _btn(bf, "تحديث", self._refresh_results, bg=BG_D).pack(
            side="left", padx=6)

    def _refresh_results(self):
        for i in self._rtree.get_children():
            self._rtree.delete(i)
        for fa in self.analyses:
            self._rtree.insert("", "end", values=[
                fa.name, fa.n_items,
                f"{fa.mean_diff:.3f}",
                f"{fa.min_diff:.3f}",
                f"{fa.max_diff:.3f}",
                f"{fa.avg_a:.3f}",
                f"{fa.avg_c:.3f}",
                f"{fa.alpha:.3f}" if not math.isnan(fa.alpha) else "—",
                f"{fa.item_corr:.3f}" if not math.isnan(fa.item_corr) else "—",
            ])

    def _open_folder(self):
        if self.source_path:
            d = str(self.source_path.parent)
            if sys.platform == "win32":
                subprocess.Popen(["explorer", d])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", d])
            else:
                subprocess.Popen(["xdg-open", d])

    # ════════════════════ GENERATE LOGIC ══════════════════════════════
    def _run(self):
        df = getattr(self, "_filtered_df", None) or self.bank_df
        if df is None:
            mb.showerror("خطأ", "حمّل بنك الأسئلة أولاً (التبويب 1).")
            return
        try:
            params = self._collect_params(df)
        except ValueError as e:
            mb.showerror("خطأ في الإعداد", str(e))
            return
        self._gbtn.config(state="disabled")
        self._pv.set(0)
        self._clr_log()
        threading.Thread(target=self._gen_thread,
                         args=(df, params), daemon=True).start()

    def _collect_params(self, df: pd.DataFrame) -> dict:
        def fi(e, n):
            try: return float(e.get().strip())
            except: raise ValueError(f"قيمة غير صالحة لـ '{n}'")
        def ii(e, n):
            try: return int(e.get().strip())
            except: raise ValueError(f"قيمة غير صالحة لـ '{n}'")

        n_forms = ii(self._nforms, "عدد النماذج")
        stage   = self._stage_var.get()
        crit    = DIFFICULTY_STAGES.get(stage, {})

        diff_range       = crit.get("Range") or (0.0, 1.0)
        diff_mean_range  = crit.get("Mean")

        # Manual override
        if "يدوي" in stage:
            try:
                parts = [float(x) for x in
                         self._manual_range_ent.get().split(",")]
                if len(parts) == 2:
                    diff_range = (parts[0], parts[1])
            except Exception:
                pass
            try:
                parts = [float(x) for x in
                         self._manual_mean_ent.get().split(",")]
                if len(parts) == 2:
                    diff_mean_range = (parts[0], parts[1])
            except Exception:
                pass

        n_students = ii(self._sim_n, "عدد الطلاب")
        seed_s = self._seed.get().strip()
        seed   = int(seed_s) if seed_s else None

        # Build assembly requests from pair requirements
        requests: list[AssemblyRequest] = []
        for row in self._pair_req_rows:
            try: cnt = int(row["ec"].get())
            except: cnt = 0
            if cnt <= 0:
                continue
            pair_key = row["pair_key"]
            الناتج  = row["en"].get()
            المؤشر  = row["em"].get()

            # Bin spec from Tab 3
            bins = []
            if self._use_manual_bins.get() and pair_key in self._bins_rows_map:
                entries = self._bins_rows_map[pair_key]
                bin_labels = [f"{i*0.1:.1f}-{(i+1)*0.1:.1f}" for i in range(10)]
                bin_counts = []
                for e in entries:
                    try: bin_counts.append(int(e.get()))
                    except: bin_counts.append(0)
                total_bc = sum(bin_counts)
                if total_bc > 0:
                    # Use LRM to scale to requested count
                    scaled = _largest_remainder(cnt, bin_counts)
                    for bi, (lbl, bc) in enumerate(zip(bin_labels, scaled)):
                        if bc > 0:
                            lo, hi = float(lbl.split("-")[0]), float(lbl.split("-")[1])
                            bins.append(BinSpec(lo, hi, bc))

            requests.append(AssemblyRequest(
                pair_key=pair_key,
                الناتج=الناتج,
                المؤشر=المؤشر,
                count=cnt,
                bins=bins,
            ))

        if not requests:
            raise ValueError(
                "لم يتم تحديد أي متطلبات أزواج.\n"
                "تأكد من تحميل البنك وتعبئة عدد الأسئلة في التبويب 2."
            )

        total_q = sum(r.count for r in requests)
        config = AssemblyConfig(
            requests=requests,
            diff_range=diff_range,
            diff_mean_range=diff_mean_range,
            min_a=0.5 if self._enable_stats.get() else None,
            w_difficulty=fi(self._wd, "ملاءمة الصعوبة"),
            w_bin=fi(self._wb,        "ملاءمة الحاوية"),
            w_usage=fi(self._wu,      "غرامة الاستخدام"),
            w_random=fi(self._wr,     "العشوائية"),
            allow_reuse=self._allow_reuse.get(),
            rng_seed=seed,
        )

        return {
            "n_forms": n_forms, "config": config,
            "n_students": n_students, "seed": seed,
            "total_q": total_q,
        }

    def _gen_thread(self, df: pd.DataFrame, params: dict):
        try:
            n     = params["n_forms"]
            cfg   = params["config"]
            names = [f"Form_{i+1}" for i in range(n)]

            self._lg(f"بدء التجميع: {n} نماذج…")
            results = assemble_forms(df, cfg, n, names)
            self._q.put(("prog", 40))

            all_warnings = []
            for res, name in zip(results, names):
                if res.warnings:
                    for w in res.warnings:
                        self._lg(f"  ⚠ {w}", "yellow")
                    all_warnings.extend(res.warnings)
                self._lg(f"  {name}: {len(res.form)} أسئلة، "
                         f"متوسط الصعوبة={res.mean_difficulty:.3f}")

            self._lg("تشغيل التحليل النفسي-التربوي (3PL)…")
            form_dfs = [r.form for r in results]
            analyses = analyse_all(form_dfs, names,
                                   n_students=params["n_students"])
            self._q.put(("prog", 75))

            for fa in analyses:
                self._lg(
                    f"  {fa.name}: ألفا={fa.alpha:.3f}, "
                    f"الارتباط={fa.item_corr:.3f}"
                )

            # Write outputs
            self._lg("كتابة ملفات الإخراج…")
            base = self.source_path.parent
            prefix = self.source_path.stem

            forms_path   = base / f"{prefix}_Forms.xlsx"
            analysis_path = base / f"{prefix}_3PL_Analysis.xlsx"
            remain_path  = base / f"{prefix}_Remaining_Questions.xlsx"

            usage_counts = results[-1].form["QuestionID"].value_counts().to_dict() \
                if results else {}

            write_forms_file(results, names, analyses,
                             forms_path, usage_counts)
            self._q.put(("prog", 88))

            if analyses:
                write_analysis_file(analyses, analysis_path)

            used_ids = set()
            for r in results:
                used_ids.update(r.form["QuestionID"].astype(str).tolist())
            write_remaining_file(df, used_ids, remain_path)

            self._q.put(("prog", 100))
            self._lg(f"\n  📄 النماذج:        {forms_path}")
            self._lg(f"  📊 التحليل 3PL:   {analysis_path}")
            self._lg(f"  📋 الأسئلة المتبقية: {remain_path}")
            self._lg("\n✓ تمّ بنجاح!", "lime")
            self._q.put(("done", analyses))

        except Exception as e:
            self._lg(f"\n✗ خطأ: {e}", "red")
            log.exception("Generation failed")
            self._q.put(("fail",))

    # ── Shared helpers ───────────────────────────────────────────────────
    def _le(self, parent, label, row, default=""):
        tk.Label(parent, text=label, bg=parent.cget("bg"),
                 fg=TXD, font=FB, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        e = _e(parent, w=16); e.insert(0, default)
        e.grid(row=row, column=1, sticky="w", pady=2)
        return e

    def _treeview(self, parent, columns, widths, row=0):
        s = ttk.Style()
        s.configure("S.Treeview.Heading", background=BG_M,
                    foreground=TXL, font=FH)
        s.configure("S.Treeview", font=FB, rowheight=22)
        s.map("S.Treeview", background=[("selected", ACC)])
        frm  = tk.Frame(parent, bg=BG_L)
        frm.grid(row=row, column=0, sticky="nsew")
        tree = ttk.Treeview(frm, columns=columns,
                            show="headings", style="S.Treeview")
        vsb = ttk.Scrollbar(frm, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        for col, w in zip(columns, widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, minwidth=40, anchor="center")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree.grid(row=0, column=0, sticky="nsew")
        frm.columnconfigure(0, weight=1); frm.rowconfigure(0, weight=1)
        return tree

    def _lg(self, msg, col=""):
        self._q.put(("log", msg, col))

    def _clr_log(self):
        self._q.put(("logclr",))

    def _append_log(self, msg, col=""):
        self._log.config(state="normal")
        tag = col or "g"
        self._log.tag_config(tag, foreground=col or "#58D68D")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait(); k = item[0]
                if k == "bank_ok":
                    self._on_bank(item[1], item[2])
                elif k == "log":
                    self._append_log(item[1], item[2] if len(item)>2 else "")
                elif k == "logclr":
                    self._log.config(state="normal")
                    self._log.delete("1.0", "end")
                    self._log.config(state="disabled")
                elif k == "prog":
                    self._pv.set(item[1])
                elif k == "done":
                    self.analyses = item[1]
                    self._gbtn.config(state="normal")
                    self._pv.set(100)
                    self._refresh_results()
                    self.nb.select(4)
                    self._status("اكتمل التوليد والتحليل.")
                    mb.showinfo(
                        "تمّ",
                        "تم توليد جميع النماذج وتحليلها.\n"
                        "تم حفظ ملفات الإخراج في نفس مجلد البنك."
                    )
                elif k == "fail":
                    self._gbtn.config(state="normal")
                    self._status("فشل التوليد — راجع السجل.")
                elif k == "err":
                    self._status(f"خطأ: {item[1]}")
                    mb.showerror("خطأ", item[1])
        except queue.Empty:
            pass
        self.after(80, self._poll)


# ════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = STGS()
    app.mainloop()
