"""
STGS — Smart Test Generation System  v3
========================================
Single-file application with two fully separate assembly modes
reached from a shared launcher screen.

MODE 1 — National / Standardised Exams
  Bank columns : Category, D, Difficulty, [Discrimination, Guessing]
  Assembly     : by Category → count
  Exam types   : Qudrat Ilmi, Qudrat Nazari, Tahsili, Cognitive Ability,
                 University Abilities

MODE 2 — Educational / Curriculum Assessments
  Bank columns : الناتج (outcome), المؤشر (indicator), المجال (domain),
                 difficulty, [تمييز, التخمين, subject, grade, language]
  Assembly     : by learning-outcome alone  OR  by (outcome, indicator) pair
  Grade stages : Science / Math / Reading × Grade 3 / 6 / 9

Shared features
  • Stratified bin sampling  →  questions SPREAD across the difficulty range
    (not clustered at the midpoint like a greedy scoring approach)
  • Mean-criterion retry loop (up to 100 attempts per form, same as originals)
  • Manual difficulty bin distribution with validation
  • Partial-fill mode: placeholder rows when questions are short
  • Per-domain bar charts embedded in every form sheet
  • Score histogram + average ICC + SEM chart per form
  • Automatic 3PL analysis (3000 students) after assembly
  • Three output files: Forms, 3PL Analysis, Remaining Questions
  • All UI text in English

Install: pip install pandas numpy openpyxl matplotlib xlsxwriter
Run    : python stgs_unified.py
"""

from __future__ import annotations

import io, logging, math, os, queue, subprocess, sys, threading
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
import numpy as np
import pandas as pd
import xlsxwriter
from openpyxl import Workbook as OPWorkbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-7s  %(message)s")
log = logging.getLogger("stgs")

# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# Mode 1 exam structures  ── names kept exactly as in the original Arabic script
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
        "MAR": 2, "MAL": 2, "MAN": 2, "MGE": 2,
        "CT1": 2, "CT2": 3, "CT3": 5,
        "VAN": 2, "VCA": 2, "VRC": 2, "VSC": 2,
    },
    "Custom": {},
}

# Mode 1 difficulty stages  ── original names: Stage1, E, M, D, Manually set
MODE1_STAGES: dict[str, dict] = {
    "Stage1":        {"Range": (0.20, 0.90), "Mean": (0.52, 0.58)},
    "E":             {"Range": (0.00, 0.45), "Mean": (0.27, 0.35)},
    "M":             {"Range": (0.40, 0.75), "Mean": (0.55, 0.61)},
    "D":             {"Range": (0.70, 1.00), "Mean": (0.80, 0.88)},
    "Manually set":  {"Range": None,          "Mean": None},
}

# Mode 2 exam structures  ── names kept exactly as in the original Arabic script
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
    "Custom": {},
}

# Mode 2 difficulty stages  ── original Arabic names from Script 2
MODE2_STAGES: dict[str, dict] = {
    "علوم صف الثالث":          {"Range": (0.20, 0.80), "Mean": (0.50, 0.52)},
    "علوم صف السادس":          {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "علوم الصف التاسع":        {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "الرياضيات الصف الثالث":   {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "الرياضيات الصف السادس":   {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "الرياضيات الصف التاسع":   {"Range": (0.20, 0.80), "Mean": (0.64, 0.68)},
    "القراءة الصف الثالث":     {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "القراءة الصف السادس":     {"Range": (0.20, 0.80), "Mean": (0.58, 0.62)},
    "القراءة الصف التاسع":     {"Range": (0.20, 0.80), "Mean": (0.48, 0.52)},
    "عام":                     {"Range": (0.20, 0.80), "Mean": (0.50, 0.52)},
    "Manually set":            {"Range": None,          "Mean": None},
}

BIN_LABELS = [f"{i*0.1:.1f}-{(i+1)*0.1:.1f}" for i in range(10)]

# Colours
# ── ETEC brand palette ────────────────────────────────────────────────────────
ETEC_NAVY    = "#2D2B6E"   # primary navy-purple  (logo text & bg)
ETEC_TEAL    = "#00AECB"   # top diamond
ETEC_GREEN   = "#3BB573"   # left diamond
ETEC_BLUE    = "#4B7EC8"   # centre diamond
ETEC_PURPLE  = "#6B4C9A"   # right diamond
ETEC_LIGHT   = "#F4F6FB"   # page background
ETEC_WHITE   = "#FFFFFF"
ETEC_BORDER  = "#D0D8EE"
ETEC_ALT     = "#EAF0FA"   # alternate row

# Legacy aliases used throughout the GUI
HDR_BG = ETEC_NAVY;  HDR_FG = ETEC_WHITE
SUB_BG = ETEC_BLUE;  ALT_BG = ETEC_ALT
ACCENT = ETEC_TEAL
CHART_C = [ETEC_BLUE, ETEC_TEAL, ETEC_GREEN, ETEC_PURPLE,
           "#E05C2A", "#F0A500", "#C0392B", "#1ABC9C"]

# ─────────────────────────────────────────────────────────────────────────────
#  COLUMN ALIAS MAPPING
# ─────────────────────────────────────────────────────────────────────────────

M1_ALIASES = {
    "QuestionID":    ["qoustionid","questionid","question_id","id","رقم"],
    "Category":      ["category","الفئة","التصنيف","الناتج","ناتج"],
    "D":             ["d","domain","المجال","مجال"],
    "Difficulty":    ["difficulty","الصعوبة","صعوبة","b"],
    "Discrimination":["تمييز","discrimination","a"],
    "Guessing":      ["التخمين","التخميين","guessing","c","مستوى التمييز"],
}

M2_ALIASES = {
    "QuestionID":  ["qoustionid","questionid","question_id","id","رقم"],
    "الناتج":      ["الناتج","outcome","learning_outcome","ناتج","category"],
    "المؤشر":      ["المؤشر","indicator","مؤشر"],
    "المجال":      ["المجال","domain","d","مجال"],
    "difficulty":  ["difficulty","الصعوبة","صعوبة","b"],
    "تمييز":       ["تمييز","discrimination","a"],
    "التخمين":     ["التخمين","التخميين","guessing","c","مستوى التمييز"],
    "subject":     ["subject","المادة"],
    "grade":       ["grade","الصف"],
    "language":    ["language","اللغة"],
}

# ─────────────────────────────────────────────────────────────────────────────
#  DATA LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def _read_raw(path: Path) -> pd.DataFrame:
    s = path.suffix.lower()
    if s in (".xlsx", ".xls", ".xlsm"):
        return pd.read_excel(path, engine="openpyxl")
    elif s == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig")
    raise ValueError(f"Unsupported file type: {s}")


def _map_cols(df: pd.DataFrame, aliases: dict) -> pd.DataFrame:
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


def load_bank_mode1(path) -> pd.DataFrame:
    path = Path(path)
    df = _map_cols(_read_raw(path), M1_ALIASES)
    if "Category" not in df.columns:
        raise ValueError("Column 'Category' not found. Ensure the file has a 'Category' column.")
    if "Difficulty" not in df.columns:
        raise ValueError("Column 'Difficulty' not found.")
    if "D" not in df.columns:
        df["D"] = "Unspecified"
    if "QuestionID" not in df.columns:
        df["QuestionID"] = [f"Q{i+1:05d}" for i in range(len(df))]
    df["Difficulty"] = pd.to_numeric(df["Difficulty"], errors="coerce").fillna(0.5).clip(0, 1)
    for col, default in [("Discrimination", 1.0), ("Guessing", 0.25)]:
        if col not in df.columns:
            df[col] = default
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)
    df = df[~df["QuestionID"].duplicated(keep="first")].reset_index(drop=True)
    log.info("Mode 1 bank loaded: %d questions, %d categories", len(df), df["Category"].nunique())
    return df


def load_bank_mode2(path) -> pd.DataFrame:
    path = Path(path)
    df = _read_raw(path)
    df.columns = [c.strip() for c in df.columns]
    df = _map_cols(df, M2_ALIASES)
    if "الناتج" not in df.columns:
        raise ValueError("Column 'الناتج' (outcome) not found. Ensure the bank has an outcome column.")
    if "difficulty" not in df.columns:
        raise ValueError("Column 'difficulty' not found.")
    if "المجال" not in df.columns:
        df["المجال"] = "Unspecified"
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
    df["_pair_key"] = df["الناتج"].astype(str) + "||" + df["المؤشر"].astype(str)
    df = df[~df["QuestionID"].duplicated(keep="first")].reset_index(drop=True)
    log.info("Mode 2 bank loaded: %d questions, %d pairs", len(df), df["_pair_key"].nunique())
    return df

# ─────────────────────────────────────────────────────────────────────────────
#  ASSEMBLY ENGINE
#  Core design: STRATIFIED BIN SAMPLING
#
#  Within each (Category/pair × bin), questions are sampled randomly from the
#  bin so difficulty is naturally spread.  An optional mean-criterion retry
#  loop (up to max_retries attempts) adjusts which random sample survives.
#  This matches the behaviour of the original scripts exactly.
# ─────────────────────────────────────────────────────────────────────────────

def _bell_bin_weights(diff_range: tuple[float, float],
                      mean_target: float,
                      n_bins: int = 10) -> list[float]:
    """
    Return bell-curve probability weights for the 10 standard bins (0.0–1.0),
    centred on `mean_target` with σ proportional to the width of `diff_range`.
    Bins outside `diff_range` receive zero weight.

    This produces distributions that look like the reference images:
      Stage 1 → broad bell centred ~0.55, non-zero across 0.1–0.9
      Easy (E) → left bell centred ~0.30, non-zero only in 0.0–0.45
      Medium (M) → narrow bell centred ~0.58, non-zero only in 0.4–0.75
      Hard (D)  → right bell centred ~0.84, non-zero only in 0.7–1.0
    """
    lo, hi  = diff_range
    span    = hi - lo
    # σ = span/3.3 gives bell-peak-to-tail ratio of ≈ 3:1, matching the
    # reference images exactly (e.g. Stage1: [1,2,2,3,2,2,1] over 7 bins)
    sigma = max(span / 3.3, 0.06)

    weights = []
    for i in range(n_bins):
        bin_lo  = round(i * 0.1, 1)
        bin_hi  = round((i + 1) * 0.1, 1)
        bin_mid = (bin_lo + bin_hi) / 2.0
        if bin_hi <= lo or bin_lo >= hi:
            weights.append(0.0)
        else:
            w = np.exp(-0.5 * ((bin_mid - mean_target) / sigma) ** 2)
            weights.append(float(w))

    if sum(weights) == 0:
        weights = [1.0 if (lo <= round((i + 0.5) * 0.1, 2) <= hi) else 0.0
                   for i in range(n_bins)]
    return weights


def _auto_bins(count: int,
               diff_range: tuple[float, float],
               mean_target: float) -> list["BinDef"]:
    """
    Automatically create bell-weighted BinDef list for a slot
    when no manual bins are specified.  Bins with allocated count=0
    are omitted.
    """
    weights = _bell_bin_weights(diff_range, mean_target)
    counts  = _largest_remainder(count, weights)
    bins: list[BinDef] = []
    for i, cnt in enumerate(counts):
        if cnt > 0:
            bins.append(BinDef(low=i*0.1, high=(i+1)*0.1, count=cnt))
    return bins


def _largest_remainder(total: int, proportions: list[float]) -> list[int]:
    if not proportions or total == 0:
        return [0] * len(proportions)
    s = sum(proportions) or 1
    floats = [total * p / s for p in proportions]
    floors = [int(f) for f in floats]
    ranked = sorted(range(len(floats)), key=lambda i: -(floats[i] - floors[i]))
    for k in range(total - sum(floors)):
        floors[ranked[k]] += 1
    return floors


@dataclass
class BinDef:
    low:   float
    high:  float
    count: int


@dataclass
class SlotDef:
    """One assembly slot: pick `count` items that match `filters`."""
    label:   str
    filters: dict          # column → value pairs applied to the bank
    count:   int
    bins:    list[BinDef] = field(default_factory=list)
    # if bins is empty → pick from the full difficulty range of this slot
    # if bins is populated → pick bin_count items from each bin window


@dataclass
class AssemblyParams:
    slots:           list[SlotDef]
    diff_range:      tuple[float, float] = (0.0, 1.0)
    diff_mean_range: Optional[tuple[float, float]] = None
    min_discrimination: Optional[float] = None   # filter تمييز / Discrimination
    allow_partial_fill: bool = True
    allow_reuse:        bool = False
    max_reuse:          int  = 3
    max_retries:        int  = 100
    rng_seed:           Optional[int] = None


def _pick_stratified_with_quota(
        pool: pd.DataFrame,
        total_count: int,
        used_ids_this_form: set,
        quota: dict,           # cat_label → max items allowed from that category
        cat_col: str,          # column in pool containing the category label
        diff_range: tuple,
        mean_target: float,
        sigma: float,
        allow_partial: bool,
        domain_label: str,
        form_number: int,
        question_usage: dict,
        dcol: str,
        rng: np.random.Generator,
        allow_reuse: bool,
        max_reuse: int,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Quantile-stratified sampling on the COMBINED D-domain pool while
    respecting hard per-category quotas.

    Algorithm
    ---------
    1. Remove already-used items (this form) and over-used items (cross-form).
    2. Sort remaining items by difficulty.
    3. Compute Gaussian CDF for each item.
    4. Walk through `total_count` equal CDF stripes from left to right.
    5. In each stripe, pick the best available item that still has quota.
       - If the stripe's best candidate has exhausted its category quota,
         try the next-closest item in the stripe.
       - If the whole stripe is exhausted by quota limits, borrow from
         the nearest available item in any adjacent stripe.
    6. Result: the domain-level distribution is a bell curve AND each
       category count is exactly respected.
    """
    import math
    qid_col  = "QuestionID"
    warnings: list[str] = []

    # Remove already-used and over-used items
    if used_ids_this_form and qid_col in pool.columns:
        pool = pool[~pool[qid_col].astype(str).isin(used_ids_this_form)].copy()
    if not allow_reuse and qid_col in pool.columns:
        overused = {q for q, c in question_usage.items() if c >= 1}
        pool = pool[~pool[qid_col].astype(str).isin(overused)].copy()
    elif allow_reuse and qid_col in pool.columns:
        overused = {q for q, c in question_usage.items() if c >= max_reuse}
        pool = pool[~pool[qid_col].astype(str).isin(overused)].copy()

    if pool.empty or total_count == 0:
        return pd.DataFrame(), warnings

    # Sort by difficulty
    diffs  = pd.to_numeric(pool[dcol], errors="coerce").fillna(mean_target)
    order  = np.argsort(diffs.to_numpy())
    sorted_pool  = pool.iloc[order].reset_index(drop=True)
    sorted_diffs = diffs.to_numpy()[order]

    # Gaussian CDF per item
    cdf_vals = np.array([
        0.5 * (1.0 + math.erf((d - mean_target) / (sigma * math.sqrt(2))))
        for d in sorted_diffs
    ])

    # Remaining quota per category (mutable copy)
    remaining_quota = dict(quota)
    chosen_positions: list[int] = []   # indices into sorted_pool

    stripe_w = 1.0 / total_count

    for k in range(total_count):
        lo_cdf     = k * stripe_w
        hi_cdf     = (k + 1) * stripe_w
        target_cdf = lo_cdf + rng.random() * stripe_w   # jittered target

        # Build candidate set: in-stripe, not yet chosen, still has quota
        def _available(idx_set):
            """Filter indices to those not yet chosen and with quota."""
            result = []
            for i in idx_set:
                if i in chosen_positions:
                    continue
                cat = str(sorted_pool.at[i, cat_col]) if cat_col in sorted_pool.columns else "__"
                if remaining_quota.get(cat, 0) > 0:
                    result.append(i)
            return result

        stripe_mask = np.where((cdf_vals >= lo_cdf) & (cdf_vals < hi_cdf))[0]
        candidates  = _available(stripe_mask)

        if not candidates:
            # Stripe exhausted by quota — expand search radius outward
            all_remaining = _available(range(len(sorted_pool)))
            if not all_remaining:
                break   # pool truly exhausted
            # Pick the item with CDF closest to stripe midpoint
            mid_cdf    = (lo_cdf + hi_cdf) / 2.0
            candidates = sorted(all_remaining,
                                 key=lambda i: abs(cdf_vals[i] - mid_cdf))

        # Among candidates, pick closest CDF to jittered target
        best = min(candidates, key=lambda i: abs(cdf_vals[i] - target_cdf))
        chosen_positions.append(best)

        # Decrement category quota
        cat = str(sorted_pool.at[best, cat_col]) if cat_col in sorted_pool.columns else "__"
        if cat in remaining_quota:
            remaining_quota[cat] -= 1

    selected = sorted_pool.iloc[chosen_positions].copy()

    # Pad with placeholders if we couldn't fill all stripes
    got = len(selected)
    if got < total_count:
        if allow_partial:
            needed = total_count - got
            # Use None for ALL columns so numeric operations never crash
            placeholder = pd.DataFrame({
                col: [None] * needed
                for col in pool.columns
            })
            selected = pd.concat([selected, placeholder], ignore_index=True)
            warnings.append(
                f"Form {form_number}: domain '{domain_label}' short by {needed} — "
                f"placeholders added.")

    return selected, warnings


def _gaussian_cdf(x: np.ndarray, mean: float, sigma: float) -> np.ndarray:
    """Gaussian CDF using the error function (no scipy needed)."""
    return 0.5 * (1.0 + np.array(
        [float(__import__('math').erf((xi - mean) / (sigma * 1.41421356)))
         for xi in x]
    ))


def _pick_stratified(pool: pd.DataFrame,
                     count: int,
                     used_ids_this_form: set,
                     diff_range: tuple,
                     mean_target: float,
                     sigma: float,
                     allow_partial: bool,
                     label: str,
                     form_number: int,
                     question_usage: dict,
                     dcol: str,
                     rng: np.random.Generator,
                     allow_reuse: bool,
                     max_reuse: int,
                     ) -> tuple[pd.DataFrame, list[str]]:
    """
    Quantile-stratified Gaussian sampling — guarantees smooth bell coverage.

    Algorithm
    ---------
    1. Compute the Gaussian CDF value for every candidate item based on its
       difficulty.
    2. Divide the CDF range [0, 1] into `count` equal-probability stripes.
    3. In each stripe, randomly pick ONE item whose CDF value falls in that
       stripe (with a small random jitter to avoid always picking the same
       boundary item).
    4. If a stripe is empty (no item in that difficulty zone), borrow the
       nearest available item from a neighbour stripe.

    Why this works for small counts
    --------------------------------
    Even with count=1 (MAL) the single item is drawn from the middle stripe
    (near the mean). With count=2 (MAN) items land in the 1/4 and 3/4 CDF
    stripes — flanking the mean symmetrically. With count=5 (MAR) items
    spread across 5 evenly-spaced quantiles, guaranteeing tail coverage.

    When many slots share the same D domain, the aggregate chart across all
    categories forms a smooth bell curve, while each slot's count is exact.
    """
    import math
    qid_col  = "QuestionID"
    warnings: list[str] = []

    # --- Remove already-used items this form ---------------------------------
    if used_ids_this_form and qid_col in pool.columns:
        pool = pool[~pool[qid_col].astype(str).isin(used_ids_this_form)].copy()

    # --- Handle exhausted / too-small pool -----------------------------------
    def _placeholders(n, cols):
        # All columns get None — prevents numeric/string crashes downstream
        return pd.DataFrame({col: [None] * n for col in cols})

    if pool.empty or count == 0:
        if count > 0 and allow_partial:
            warnings.append(f"Form {form_number}: no items for '{label}'.")
            return _placeholders(count, pool.columns if not pool.empty
                                 else [dcol, "QuestionID"]), warnings
        return pd.DataFrame(), warnings

    # --- Build working pool (optionally include reuse candidates) ------------
    working = pool.copy()
    if not allow_reuse:
        pass  # pool is already filtered upstream
    else:
        working = working[
            working[qid_col].astype(str).map(
                lambda q: question_usage.get(q, 0) < max_reuse)
        ].copy()
    if working.empty:
        working = pool.copy()  # last resort

    # --- If pool too small, take all + pad -----------------------------------
    if len(working) < count:
        if allow_partial:
            needed = count - len(working)
            result = pd.concat(
                [working, _placeholders(needed, working.columns)],
                ignore_index=True)
            warnings.append(
                f"Form {form_number}: shortage of {needed} for '{label}' "
                f"— placeholders added.")
            return result, warnings
        else:
            return _placeholders(count, working.columns), warnings

    # --- Compute per-item CDF values based on difficulty ---------------------
    diffs = pd.to_numeric(working[dcol], errors="coerce").fillna(mean_target)
    cdf_vals = np.array([
        0.5 * (1.0 + math.erf((d - mean_target) / (sigma * math.sqrt(2))))
        for d in diffs
    ])

    # --- Sort by CDF (= sort by difficulty) ----------------------------------
    order    = np.argsort(cdf_vals)
    sorted_w = working.iloc[order].reset_index(drop=True)
    sorted_c = cdf_vals[order]

    # --- Divide [0,1] into `count` equal stripes and pick one per stripe -----
    # Add tiny random offset within each stripe (Latin-Hypercube style)
    stripe_w = 1.0 / count
    chosen_indices: list[int] = []

    for k in range(count):
        lo_cdf = k * stripe_w
        hi_cdf = (k + 1) * stripe_w
        # jitter: shift the target point randomly within the stripe
        target_cdf = lo_cdf + rng.random() * stripe_w

        # Find all candidates in this stripe
        mask = (sorted_c >= lo_cdf) & (sorted_c < hi_cdf)
        # Exclude already chosen
        mask &= ~np.isin(np.arange(len(sorted_w)), chosen_indices)

        if mask.any():
            candidates = np.where(mask)[0]
            # Among stripe candidates, pick the one whose CDF is closest
            # to the jittered target (deterministic within jitter)
            best = candidates[np.argmin(np.abs(sorted_c[candidates] - target_cdf))]
            chosen_indices.append(int(best))
        else:
            # Stripe empty — find the nearest unused item by CDF distance
            unused_mask = ~np.isin(np.arange(len(sorted_w)), chosen_indices)
            if unused_mask.any():
                unused_idx   = np.where(unused_mask)[0]
                cdf_target_m = (lo_cdf + hi_cdf) / 2.0
                nearest      = unused_idx[
                    np.argmin(np.abs(sorted_c[unused_idx] - cdf_target_m))
                ]
                chosen_indices.append(int(nearest))

    selected = sorted_w.iloc[chosen_indices].copy()
    return selected, warnings


def _pick_random(pool: pd.DataFrame,
                 count: int,
                 used_ids_this_form: set,   # QuestionID strings already in this form
                 allow_partial: bool,
                 label: str,
                 form_number: int,
                 question_usage: dict,       # qid → cross-form use count
                 dcol: str,
                 rng: np.random.Generator,
                 allow_reuse: bool,
                 max_reuse: int,
                 ) -> tuple[pd.DataFrame, list[str]]:
    """
    Pick `count` items from `pool`.

    Uniqueness rules
    ----------------
    • A question is NEVER selected twice in the same form (guaranteed by
      used_ids_this_form — keyed on QuestionID, so bin fallbacks cannot
      re-pick something already chosen earlier in the same form).
    • Across forms, reuse is controlled by allow_reuse / max_reuse:
        allow_reuse=False  → question used in any previous form is excluded
        allow_reuse=True   → question may appear in up to max_reuse (2) forms
    """
    qid_col = "QuestionID"
    warnings: list[str] = []

    # Exclude items already selected this form (by QuestionID — the true key)
    if used_ids_this_form and qid_col in pool.columns:
        pool = pool[~pool[qid_col].astype(str).isin(used_ids_this_form)].copy()

    if len(pool) >= count:
        idx = rng.choice(len(pool), size=count, replace=False)
        selected = pool.iloc[idx].copy()
    elif allow_partial:
        selected = pool.copy()
        needed = count - len(selected)
        if needed > 0:
            # All None — never put strings in numeric columns
            placeholder = pd.DataFrame(
                {col: [None] * needed for col in pool.columns})
            selected = pd.concat([selected, placeholder], ignore_index=True)
            warnings.append(
                f"Form {form_number}: shortage of {needed} for '{label}' "
                f"— placeholder rows added.")
    else:
        # Try reuse pool: items whose cross-form count < max_reuse
        if allow_reuse and qid_col in pool.columns:
            reuse_pool = pool[
                pool[qid_col].astype(str).map(
                    lambda q: question_usage.get(q, 0) < max_reuse
                )
            ].copy()
        else:
            reuse_pool = pool.copy()

        if len(reuse_pool) >= count:
            idx = rng.choice(len(reuse_pool), size=count, replace=False)
            selected = reuse_pool.iloc[idx].copy()
        else:
            # All None — never put strings in numeric columns
            selected = pd.DataFrame(
                {col: [None] * count for col in pool.columns})
            warnings.append(
                f"Form {form_number}: not enough items for '{label}' "
                f"— {count} placeholder rows added.")

    return selected, warnings


class AssemblyEngine:
    """
    Stratified bell-curve bin sampling engine.

    Uniqueness guarantee
    --------------------
    Each form maintains an `intra_form_used_ids` set of QuestionIDs.
    This set is passed into every _pick_random call, so it is impossible
    for the same question to appear twice in a single form — even when
    bin fallback widens the pool to the full slot or global bank.

    Cross-form reuse
    ----------------
    Controlled by AssemblyParams.allow_reuse and .max_reuse (default 2).
    When allow_reuse=True a question may appear in at most max_reuse
    different forms but NEVER more than once within any single form.
    """

    def __init__(self, bank: pd.DataFrame, params: AssemblyParams,
                 dcol: str,
                 usage: Optional[dict] = None):
        self.bank   = bank.copy().reset_index(drop=True)
        self.params = params
        self.dcol   = dcol
        self.rng    = np.random.default_rng(params.rng_seed)
        # usage[QuestionID_str] = number of forms it has been used in
        self.question_usage: dict[str, int] = usage or {}

    # ── public ────────────────────────────────────────────────────────────────
    def assemble_one_form(self, form_number: int
                          ) -> tuple[pd.DataFrame, float, list[str]]:
        """
        Returns (form_df, mean_difficulty, warnings).

        Distribution strategy
        ---------------------
        When manual bins are NOT set we use form-level bell-curve orchestration:
          1. Compute the bell-curve bin allocations for the TOTAL question count.
          2. For each bin, draw from the combined pool of ALL slots together.
             This ensures every bin gets filled, even when individual slots
             each contribute only 1–2 questions.
          3. After selecting from each bin, tag each row with its originating
             slot label (for export) using a best-match approach.

        When manual bins ARE set on at least one slot, fall back to the
        original per-slot approach so the manual distribution is honoured.
        """
        p = self.params
        intra_form_used_ids: set[str] = set()
        form_parts: list[pd.DataFrame] = []
        all_warnings: list[str] = []

        mean_t = (float(np.mean(p.diff_mean_range))
                  if p.diff_mean_range
                  else (p.diff_range[0] + p.diff_range[1]) / 2.0)
        lo, hi = p.diff_range
        sigma  = max((hi - lo) / 3.3, 0.06)
        use_manual = any(s.bins for s in p.slots)

        if use_manual:
            # ── Manual bins: per-slot, draw from each bin range exactly ──────
            for slot in p.slots:
                pool = self._pool_for_slot(slot, intra_form_used_ids)
                for bdef in slot.bins:
                    bin_pool = pool[
                        (pool[self.dcol] >= bdef.low) &
                        (pool[self.dcol] <  bdef.high)
                    ].copy()
                    if bin_pool.empty:
                        bin_pool = pool.copy()
                    sel, warns = _pick_random(
                        bin_pool, bdef.count, intra_form_used_ids,
                        p.allow_partial_fill,
                        f"{slot.label} [{bdef.low:.1f}-{bdef.high:.1f}]",
                        form_number, self.question_usage,
                        self.dcol, self.rng, p.allow_reuse, p.max_reuse,
                    )
                    sel["_form"] = form_number
                    form_parts.append(sel)
                    all_warnings.extend(warns)
                    if "QuestionID" in sel.columns:
                        intra_form_used_ids.update(
                            sel["QuestionID"].dropna().astype(str))
        else:
            # ── D-domain stratified bell sampling ─────────────────────────────
            #
            # The bell curve belongs to the D domain as a whole:
            #   • Group all slots by their D value
            #   • Build ONE combined D-pool (all categories in that domain)
            #   • Apply quantile-stratified sampling on the combined pool
            #     for the total domain question count
            #   • Hard per-category quotas (MAR≤5, MAL≤1, …) are enforced
            #     during item selection so exact counts are always met
            #
            domain_col = "D" if "D" in self.bank.columns else "المجال"

            # --- Step 1: discover D value for each slot ----------------------
            slot_to_domain: dict[str, str] = {}
            for slot in p.slots:
                sp = self._pool_for_slot(slot, intra_form_used_ids)
                if not sp.empty and domain_col in sp.columns:
                    dom = str(sp[domain_col].mode().iloc[0])
                else:
                    dom = "__no_domain__"
                slot_to_domain[slot.label] = dom

            # --- Step 2: group slots by domain --------------------------------
            domain_to_slots: dict[str, list] = {}
            for slot in p.slots:
                d = slot_to_domain[slot.label]
                domain_to_slots.setdefault(d, []).append(slot)

            # --- Step 3: for each domain apply one stratified bell curve ------
            for domain, d_slots in domain_to_slots.items():

                # Remaining quota per category (starts at the target count)
                quota: dict[str, int] = {s.label: s.count for s in d_slots}

                # Combined pool for this domain: all categories merged
                d_pools = []
                for slot in d_slots:
                    sp = self._pool_for_slot(slot, intra_form_used_ids)
                    sp = sp.copy()
                    sp["_cat_label"] = slot.label
                    d_pools.append(sp)
                d_pool = pd.concat(d_pools, ignore_index=True) \
                         if d_pools else pd.DataFrame()

                if d_pool.empty:
                    continue

                # Total items needed from this domain
                total_needed = sum(quota.values())

                # Apply stratified sampling on the combined D pool,
                # honouring per-category hard caps
                sel, warns = _pick_stratified_with_quota(
                    d_pool, total_needed, intra_form_used_ids,
                    quota, "_cat_label",
                    p.diff_range, mean_t, sigma,
                    p.allow_partial_fill, domain, form_number,
                    self.question_usage, self.dcol, self.rng,
                    p.allow_reuse, p.max_reuse,
                )
                sel["_form"] = form_number
                form_parts.append(sel)
                all_warnings.extend(warns)
                if "QuestionID" in sel.columns:
                    intra_form_used_ids.update(
                        sel["QuestionID"].dropna().astype(str))

        form = pd.concat(form_parts, ignore_index=True) if form_parts else pd.DataFrame()
        # Drop helper columns before returning
        form = form.drop(columns=["_slot_label"], errors="ignore")
        difficulties = pd.to_numeric(form[self.dcol], errors="coerce").dropna().tolist()
        mean_d = float(np.mean(difficulties)) if difficulties else 0.0
        return form, mean_d, all_warnings

    # ── helpers ───────────────────────────────────────────────────────────────
    def _pool_for_slot(self, slot: SlotDef,
                       intra_form_used_ids: set[str]) -> pd.DataFrame:
        """
        Build the candidate pool for a slot.

        Reuse-as-last-resort policy
        ----------------------------
        Reuse (questions that appeared in a previous form) is NEVER the first
        choice — regardless of whether the 'Allow reuse' checkbox is ticked.

        Priority order:
          1. Fresh questions (never used in any form)          ← always tried first
          2. Once-used questions (only when allow_reuse=True
             AND fresh pool is too small)                      ← last resort
          3. Least-used first (sorted by use count ascending)  ← ordering for (2)

        Questions already placed in THIS form are always excluded (hard rule).
        Questions used >= max_reuse times are always excluded (hard cap).
        """
        p    = self.params
        pool = self.bank.copy()

        # Apply slot column filters (Category / الناتج / المؤشر)
        for col, val in slot.filters.items():
            if col in pool.columns and val not in (None, "", "nan"):
                pool = pool[pool[col].astype(str) == str(val)]

        # Discrimination filter
        disc_col = "Discrimination" if "Discrimination" in pool.columns else "تمييز"
        if p.min_discrimination is not None and disc_col in pool.columns:
            pool = pool[pool[disc_col] >= p.min_discrimination]

        # Global difficulty range
        lo, hi = p.diff_range
        pool = pool[(pool[self.dcol] >= lo) & (pool[self.dcol] <= hi)]

        # Hard cap: always exclude questions used >= max_reuse times
        if "QuestionID" in pool.columns:
            hard_cap = {q for q, c in self.question_usage.items()
                        if c >= p.max_reuse}
            pool = pool[~pool["QuestionID"].astype(str).isin(hard_cap)]

        # Hard rule: never duplicate within the current form
        if intra_form_used_ids and "QuestionID" in pool.columns:
            pool = pool[~pool["QuestionID"].astype(str).isin(intra_form_used_ids)]

        if pool.empty:
            return pool.reset_index(drop=True)

        # ── Priority 1: FRESH questions (used in 0 previous forms) ────────────
        fresh_ids = {q for q, c in self.question_usage.items() if c >= 1}
        if "QuestionID" in pool.columns:
            fresh_pool = pool[~pool["QuestionID"].astype(str).isin(fresh_ids)]
        else:
            fresh_pool = pool.copy()

        if len(fresh_pool) >= slot.count:
            # Enough fresh questions — return fresh only
            return fresh_pool.reset_index(drop=True)

        # ── Priority 2: Allow reuse only if checkbox is ON and fresh exhausted ─
        if not p.allow_reuse:
            # Reuse not permitted — return whatever fresh questions exist
            # (shortfall will be handled by placeholder logic upstream)
            return fresh_pool.reset_index(drop=True)

        # Reuse IS allowed but only as last resort:
        # Return the full pool sorted by usage count ascending
        # (fresh first, then least-used, then more-used)
        pool = pool.copy()
        if "QuestionID" in pool.columns:
            pool["_usage"] = pool["QuestionID"].astype(str).map(
                lambda q: self.question_usage.get(q, 0))
            pool = pool.sort_values("_usage").drop(columns=["_usage"])

        return pool.reset_index(drop=True)


def assemble_forms(bank: pd.DataFrame,
                   params: AssemblyParams,
                   dcol: str,
                   n_forms: int,
                   form_names: list[str]
                   ) -> list[tuple[pd.DataFrame, float, list[str]]]:
    """
    Assemble n_forms sharing usage tracking so questions spread.
    Uses a retry loop (up to params.max_retries) per form to satisfy
    the mean criterion — mirrors the original 100-retry logic.
    """
    # usage[QuestionID_str] = how many forms it has appeared in so far
    usage: dict[str, int] = {}
    results: list[tuple[pd.DataFrame, float, list[str]]] = []
    mn_range = params.diff_mean_range

    for fidx, name in enumerate(form_names):
        engine = AssemblyEngine(bank, params, dcol, dict(usage))
        best_form, best_mean, best_warns = engine.assemble_one_form(fidx + 1)
        success = (mn_range is None or
                   (mn_range[0] <= best_mean <= mn_range[1]))

        if not success:
            for attempt in range(1, params.max_retries):
                seed = (params.rng_seed or 0) + fidx * 1000 + attempt
                p2 = AssemblyParams(**{
                    **params.__dict__,
                    "rng_seed": seed,
                    "diff_mean_range": None,
                })
                eng2 = AssemblyEngine(bank, p2, dcol, dict(usage))
                form2, mean2, w2 = eng2.assemble_one_form(fidx + 1)
                if mn_range[0] <= mean2 <= mn_range[1]:
                    best_form, best_mean, best_warns = form2, mean2, w2
                    success = True
                    break
            if not success:
                best_warns.append(
                    f"{name}: mean criterion ({mn_range[0]:.2f}–{mn_range[1]:.2f}) "
                    f"not met after {params.max_retries} retries "
                    f"(achieved {best_mean:.3f})."
                )

        # Update cross-form usage by QuestionID (the only reliable unique key)
        if "QuestionID" in best_form.columns:
            for qid in best_form["QuestionID"].dropna().astype(str):
                usage[qid] = usage.get(qid, 0) + 1

        log.info("%s: %d items, mean_diff=%.3f%s",
                 name, len(best_form), best_mean,
                 " [mean OK]" if success else " [mean MISSED]")
        results.append((best_form, best_mean, best_warns))

    return results

# ─────────────────────────────────────────────────────────────────────────────
#  3PL PSYCHOMETRIC ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

def _p3pl(theta, a, b, c):
    theta = np.asarray(theta, float).reshape(-1, 1)
    a = np.asarray(a, float).reshape(1, -1)
    b = np.asarray(b, float).reshape(1, -1)
    c = np.asarray(c, float).reshape(1, -1)
    return np.clip(c + (1 - c) / (1 + np.exp(-a * (theta - b))), 1e-6, 1 - 1e-6)


def _item_info(theta, a, b, c):
    P  = _p3pl(theta, a, b, c)
    Q  = 1.0 - P
    a2 = np.asarray(a, float).reshape(1, -1)
    c2 = np.asarray(c, float).reshape(1, -1)
    return np.clip(a2**2 * ((P - c2) / (1 - c2 + 1e-9))**2 * (Q / (P + 1e-9)), 0, 100)


def _cronbach(X):
    K = X.shape[0]
    if K < 2: return float("nan")
    tv = X.sum(axis=0).var(ddof=1)
    if tv == 0: return float("nan")
    return float(np.clip((K / (K - 1)) * (1 - X.var(axis=1, ddof=1).sum() / tv), -1, 1))


def _item_corr(X):
    try:
        cm = np.corrcoef(X)
        return float(np.nanmean(cm[np.triu_indices_from(cm, k=1)]))
    except Exception:
        return float("nan")


@dataclass
class FormAnalysis:
    name:      str
    n_items:   int
    mean_diff: float
    min_diff:  float
    max_diff:  float
    avg_a:     float
    avg_c:     float
    alpha:     float
    item_corr: float
    theta_grid: np.ndarray   # (61,)
    icc:        np.ndarray   # (61, K)
    tcc:        np.ndarray   # (61,)
    tif:        np.ndarray   # (61,)
    sem:        np.ndarray   # (61,)
    scores:     np.ndarray   # (n_students,)


def analyse_form(form_df: pd.DataFrame, name: str,
                 dcol: str, acol: str, ccol: str,
                 n_students: int = 3000,
                 rng_seed: Optional[int] = None) -> Optional[FormAnalysis]:
    if form_df.empty:
        return None
    df = form_df.copy()
    df["_b"] = pd.to_numeric(df.get(dcol, 0.5), errors="coerce").fillna(0.5).clip(0, 1)
    df["_a"] = pd.to_numeric(df.get(acol, 1.0), errors="coerce").fillna(1.0)
    df["_c"] = pd.to_numeric(df.get(ccol, 0.25), errors="coerce").fillna(0.25)
    df = df.dropna(subset=["_b"])
    if df.empty:
        return None
    b = df["_b"].to_numpy(float)
    a = df["_a"].to_numpy(float)
    c = df["_c"].to_numpy(float)
    K = len(df)
    theta_grid = np.linspace(-3, 3, 61)
    icc = _p3pl(theta_grid, a, b, c)           # (61, K)
    tcc = icc.sum(axis=1)
    tif = _item_info(theta_grid, a, b, c).sum(axis=1)
    sem = 1.0 / np.sqrt(np.clip(tif, 1e-8, None))
    rng     = np.random.default_rng(rng_seed)
    theta_s = rng.normal(0, 1, n_students)
    P_s     = _p3pl(theta_s, a, b, c)          # (n_students, K)
    X       = (rng.random((n_students, K)) < P_s).astype(np.float32).T  # (K, n_students)
    scores  = X.sum(axis=0)
    return FormAnalysis(
        name=name, n_items=K,
        mean_diff=float(b.mean()), min_diff=float(b.min()), max_diff=float(b.max()),
        avg_a=float(a.mean()), avg_c=float(c.mean()),
        alpha=_cronbach(X), item_corr=_item_corr(X),
        theta_grid=theta_grid, icc=icc, tcc=tcc, tif=tif, sem=sem, scores=scores,
    )

# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

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
    f["red"]  = wb.add_format({"font_color":"#C00000","bold":True,"border":1,
                                "bg_color":"#FFE4E4"})
    f["warn"] = wb.add_format({"bg_color":"#FFEB9C","border":1,"font_color":"#9C5700"})
    return f


# ── per-domain difficulty bar chart (replicates add_subdomain_chart()) ────────
def _write_domain_chart(wb, ws, domain_df: pd.DataFrame, dcol: str,
                        domain_label: str, sheet_name: str,
                        col_offset: int, row_offset: int) -> int:
    """Write bin table + bar chart for one domain. Returns next free row."""
    counts = [0] * 10
    for d in domain_df[dcol].dropna():
        idx = min(int(float(d) * 10), 9)
        counts[idx] += 1

    fmts = _wb_fmts(wb)
    ws.write(row_offset, col_offset,
             f"{domain_label} — Difficulty Range", fmts["shdr"])
    ws.write(row_offset, col_offset + 1, "Count", fmts["shdr"])
    for i, (lbl, cnt) in enumerate(zip(BIN_LABELS, counts)):
        ws.write(row_offset + 1 + i, col_offset,     lbl)
        ws.write(row_offset + 1 + i, col_offset + 1, cnt)

    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name":       "Q Distribution",
        "categories": [sheet_name, row_offset+1, col_offset,   row_offset+10, col_offset],
        "values":     [sheet_name, row_offset+1, col_offset+1, row_offset+10, col_offset+1],
        "data_labels": {"value": True},
        "fill":        {"color": "#2E75B6"},
    })
    chart.set_title({"name": f"{domain_label} — Difficulty Distribution"})
    chart.set_x_axis({"name": "Difficulty Range"})
    chart.set_y_axis({"name": "Count", "min": 0})
    chart.set_legend({"position": "right"})
    chart.set_size({"width": 480, "height": 288})
    ws.insert_chart(row_offset + 12, col_offset, chart)
    return row_offset + 30


def _embed_3pl_charts(wb, ws, fa: FormAnalysis, img_row: int):
    """Embed score histogram, average ICC, SEM charts into the worksheet."""
    # Score distribution
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.hist(fa.scores, bins=15, color="#2E75B6", edgecolor="white", alpha=0.88)
    ax.set(xlabel="Total Score", ylabel="Number of Students",
           title="Simulated Score Distribution")
    ax.grid(alpha=0.3, axis="y"); fig.tight_layout()
    ws.insert_image(img_row, 0, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)

    # Average ICC
    avg_icc = fa.icc.mean(axis=1)
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot(fa.theta_grid, avg_icc, color="#ED7D31", lw=2)
    ax.set(xlabel="Ability (θ)", ylabel="P(Correct)",
           title="Item Characteristic Curve (ICC)", xlim=(-3, 3), ylim=(0, 1.05))
    ax.grid(alpha=0.3); fig.tight_layout()
    ws.insert_image(img_row, 7, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)

    # SEM
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot(fa.theta_grid, fa.sem, color="#70AD47", lw=2)
    ax.set(xlabel="Ability (θ)", ylabel="SEM",
           title="Standard Error of Measurement (SEM)", xlim=(-3, 3))
    ax.grid(alpha=0.3); fig.tight_layout()
    ws.insert_image(img_row + 20, 0, "", {"image_data": _fig_bytes(fig)})
    plt.close(fig)


M1_EXPORT_COLS = [
    ("QuestionID", 18, False), ("Category", 22, False),
    ("D", 16, False),
    ("Difficulty", 12, True), ("Discrimination", 14, True), ("Guessing", 12, True),
]
M2_EXPORT_COLS = [
    ("QuestionID", 18, False), ("الناتج", 22, False),
    ("المؤشر", 22, False), ("المجال", 18, False),
    ("difficulty", 12, True), ("تمييز", 12, True), ("التخمين", 12, True),
]


def write_forms_workbook(
    results: list[tuple[pd.DataFrame, float, list[str]]],
    form_names: list[str],
    analyses: list[Optional[FormAnalysis]],
    dcol: str,
    domain_col: str,          # "D" for mode1, "المجال" for mode2
    acol: str,
    ccol: str,
    question_usage: dict,     # original_index → total use count
    out_path: Path,
    mode: int,
):
    wb   = xlsxwriter.Workbook(str(out_path))
    fmts = _wb_fmts(wb)
    fa_map = {fa.name: fa for fa in analyses if fa}
    export_cols = M1_EXPORT_COLS if mode == 1 else M2_EXPORT_COLS

    for (form_df, mean_d, warns), name in zip(results, form_names):
        ws = wb.add_worksheet(name[:31])
        df = form_df.reset_index(drop=True)
        fa = fa_map.get(name)

        # ── Summary header ──────────────────────────────────────────────────
        ws.merge_range("A1:D1", name, fmts["ttl"])
        ws.set_row(0, 22)

        # Column widths for the summary block
        ws.set_column(0, 0, 30)   # Metric label
        ws.set_column(1, 1, 14)   # Overall value

        # ── Overall metrics (left block, columns A-B) ────────────────────────
        overall_metrics = [
            ("Number of Items",           len(df),                             False),
            ("Mean Difficulty",           df[dcol].mean() if dcol in df else "—", True),
            ("Min Difficulty",            df[dcol].min()  if dcol in df else "—", True),
            ("Max Difficulty",            df[dcol].max()  if dcol in df else "—", True),
            ("Avg Discrimination (a)",    df[acol].mean() if acol in df else "—", True),
            ("Simulated Cronbach Alpha",  fa.alpha     if fa else "—",             True),
            ("Simulated Item Correlation",fa.item_corr if fa else "—",             True),
        ]
        ws.write(2, 0, "Metric",        fmts["shdr"])
        ws.write(2, 1, "Overall",       fmts["shdr"])
        for ri, (lbl, val, is_f) in enumerate(overall_metrics):
            r = 3 + ri
            ws.write(r, 0, lbl, fmts["mlbl"])
            if isinstance(val, str):
                ws.write(r, 1, val, fmts["dat"])
            elif is_f and val is not None:
                try:    ws.write_number(r, 1, float(val), fmts["mval"])
                except: ws.write(r, 1, str(val), fmts["dat"])
            else:
                ws.write_number(r, 1, int(val) if val else 0, fmts["mvi"])

        # ── Per-D-domain stats (columns C onwards, one column per domain) ────
        if domain_col in df.columns:
            domains = [d for d in df[domain_col].dropna().unique()
                       if str(d).strip()]
            dom_start_col = 2          # column C = index 2
            for di, dom in enumerate(domains):
                col_idx = dom_start_col + di
                col_w   = 14
                ws.set_column(col_idx, col_idx, col_w)
                sub = df[df[domain_col] == dom]
                sub_diff = pd.to_numeric(sub[dcol], errors="coerce").dropna() \
                           if dcol in sub.columns else pd.Series(dtype=float)

                # Domain column header
                ws.write(2, col_idx, f"Domain: {dom}", fmts["shdr"])

                # Rows: N items, Mean, Min, Max for this domain, rest blank
                domain_metrics = [
                    len(sub),
                    float(sub_diff.mean()) if not sub_diff.empty else None,
                    float(sub_diff.min())  if not sub_diff.empty else None,
                    float(sub_diff.max())  if not sub_diff.empty else None,
                    None, None, None,  # Discrimination / Alpha / Corr not per-domain
                ]
                for ri, val in enumerate(domain_metrics):
                    r = 3 + ri
                    if val is None:
                        ws.write(r, col_idx, "—", fmts["dat"])
                    elif ri == 0:
                        ws.write_number(r, col_idx, int(val), fmts["mvi"])
                    else:
                        ws.write_number(r, col_idx, val, fmts["mval"])

        # ── Warnings ────────────────────────────────────────────────────────
        warn_row = 11
        if warns:
            ws.write(warn_row, 0, "⚠ Warnings", fmts["warn"])
            for wi, w in enumerate(warns[:6]):
                ws.write(warn_row + 1 + wi, 0, w, fmts["warn"])
            warn_row += len(warns[:6]) + 2
        tbl = warn_row + 1

        # ── Item table ──────────────────────────────────────────────────────
        for ci, (cn, cw, _) in enumerate(export_cols):
            ws.write(tbl, ci, cn, fmts["shdr"])
            ws.set_column(ci, ci, cw)

        for ri, (_, row) in enumerate(df.iterrows()):
            qid    = str(row.get("QuestionID", ""))
            reused = question_usage.get(qid, 0) > 1
            for ci, (cn, _, is_n) in enumerate(export_cols):
                val = row.get(cn, "")
                if is_n:
                    fmt = fmts["nu2"] if ri % 2 else fmts["num"]
                    try:    ws.write_number(tbl+1+ri, ci, float(val), fmt)
                    except: ws.write(tbl+1+ri, ci, str(val), fmt)
                else:
                    fmt = (fmts["red"] if reused
                           else (fmts["da2"] if ri % 2 else fmts["dat"]))
                    ws.write(tbl+1+ri, ci, str(val) if val is not None else "", fmt)

        ws.freeze_panes(tbl + 1, 0)
        ws.autofilter(tbl, 0, tbl + len(df), len(export_cols) - 1)

        # ── Per-domain difficulty bar charts (replicates add_subdomain_chart) ──
        chart_col = len(export_cols) + 2
        chart_row = tbl
        if domain_col in df.columns:
            for dom in df[domain_col].dropna().unique():
                sub = df[df[domain_col] == dom]
                if dcol in sub.columns:
                    chart_row = _write_domain_chart(
                        wb, ws, sub, dcol, str(dom), name[:31],
                        chart_col, chart_row)

        # ── 3PL charts (score hist, ICC, SEM) ───────────────────────────────
        if fa:
            img_row = tbl + len(df) + 4
            _embed_3pl_charts(wb, ws, fa, img_row)

    wb.close()
    log.info("Forms workbook → %s", out_path)


def write_analysis_workbook(analyses: list[FormAnalysis], out_path: Path):
    wb = OPWorkbook(); wb.remove(wb.active)

    # Summary sheet
    ws = wb.create_sheet("Summary")
    ws.append(["Form","N Items","Avg a","Avg b","Avg c","Mean Diff",
               "Min Diff","Max Diff","Cronbach Alpha","Item Corr."])
    for fa in analyses:
        ws.append([fa.name, fa.n_items,
                   round(fa.avg_a,4), round(fa.mean_diff,4), round(fa.avg_c,4),
                   round(fa.mean_diff,4), round(fa.min_diff,4), round(fa.max_diff,4),
                   round(fa.alpha,4) if not math.isnan(fa.alpha) else "N/A",
                   round(fa.item_corr,4) if not math.isnan(fa.item_corr) else "N/A"])
    # Overall average row
    avgs = {f: float(np.nanmean([getattr(fa,f) for fa in analyses]))
            for f in ("avg_a","avg_c","mean_diff","min_diff","max_diff","alpha","item_corr")}
    ws.append(["Overall Average","",
               round(avgs["avg_a"],4),"","",
               round(avgs["mean_diff"],4),round(avgs["min_diff"],4),round(avgs["max_diff"],4),
               round(avgs["alpha"],4),round(avgs["item_corr"],4)])

    # ICC per form
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

    # TCC comparison
    ws_tcc = wb.create_sheet("TCC_Comparison")
    ws_tcc.append(["Theta"] + [fa.name for fa in analyses])
    for i, th in enumerate(analyses[0].theta_grid):
        ws_tcc.append([round(th,3)] + [round(float(fa.tcc[i]),4) for fa in analyses])
    n = len(analyses[0].theta_grid)
    chart = LineChart(); chart.title = "Test Characteristic Curves (TCC)"
    chart.y_axis.title = "Expected Score"; chart.x_axis.title = "Theta"; chart.style = 10
    dr = Reference(ws_tcc, min_col=2, min_row=1, max_col=1+len(analyses), max_row=1+n)
    cr = Reference(ws_tcc, min_col=1, min_row=2, max_row=1+n)
    chart.add_data(dr, titles_from_data=True); chart.set_categories(cr)
    ws_tcc.add_chart(chart, f"{get_column_letter(len(analyses)+3)}2")

    # TIF comparison
    ws_tif = wb.create_sheet("TIF_Comparison")
    ws_tif.append(["Theta"] + [fa.name for fa in analyses])
    for i, th in enumerate(analyses[0].theta_grid):
        ws_tif.append([round(th,3)] + [round(float(fa.tif[i]),4) for fa in analyses])
    chart = LineChart(); chart.title = "Test Information Function (TIF)"
    chart.y_axis.title = "Information"; chart.x_axis.title = "Theta"; chart.style = 10
    dr = Reference(ws_tif, min_col=2, min_row=1, max_col=1+len(analyses), max_row=1+n)
    cr = Reference(ws_tif, min_col=1, min_row=2, max_row=1+n)
    chart.add_data(dr, titles_from_data=True); chart.set_categories(cr)
    ws_tif.add_chart(chart, f"{get_column_letter(len(analyses)+3)}2")

    wb.save(str(out_path))
    log.info("3PL analysis → %s", out_path)

# ─────────────────────────────────────────────────────────────────────────────
#  GUI SHARED HELPERS
# ─────────────────────────────────────────────────────────────────────────────

FT = ("Segoe UI", 14, "bold"); FH = ("Segoe UI", 11, "bold")
FB = ("Segoe UI", 10);          FM = ("Consolas", 9)
BG_D = ETEC_NAVY;   BG_M = ETEC_BLUE;  BG_L = ETEC_LIGHT
TXD = ETEC_NAVY;    TXL = ETEC_WHITE


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
    frm  = tk.Frame(parent, bg=BG_L); frm.pack(fill="both", expand=True)
    tree = ttk.Treeview(frm, columns=columns, show="headings", style="G.Treeview")
    vsb  = ttk.Scrollbar(frm, orient="vertical",   command=tree.yview)
    hsb  = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    for col, w in zip(columns, widths):
        tree.heading(col, text=col); tree.column(col, width=w, minwidth=40, anchor="center")
    vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
    tree.grid(row=0, column=0, sticky="nsew")
    frm.columnconfigure(0, weight=1); frm.rowconfigure(0, weight=1)
    return tree


def _log_widget(parent):
    frm = tk.Frame(parent, bg=BG_L); frm.pack(fill="both", expand=True)
    box = tk.Text(frm, bg="#0F1A2E", fg=ETEC_TEAL, font=FM,
                  state="disabled", relief="flat", height=18)
    sb  = ttk.Scrollbar(frm, command=box.yview)
    box.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y"); box.pack(fill="both", expand=True)
    return box


def _append_log(box, msg, col=""):
    box.config(state="normal")
    tag = col or "g"; box.tag_config(tag, foreground=col or ETEC_TEAL)
    box.insert("end", msg + "\n", tag); box.see("end")
    box.config(state="disabled")

# ─────────────────────────────────────────────────────────────────────────────
#  LAUNCHER WINDOW
# ─────────────────────────────────────────────────────────────────────────────

def _draw_etec_logo(canvas: tk.Canvas, x: int, y: int, scale: float = 1.0):
    """
    Draw the ETEC diamond-leaf logo programmatically on a Tkinter Canvas.
    The logo consists of 8 diamond shapes arranged in a leaf/flower pattern.
      x, y  = top-left origin of the bounding box
      scale = size multiplier (1.0 → approx 80×70 px)
    """
    s = scale
    # Diamond helper: draw a rotated-square diamond
    def diamond(cx, cy, w, h, fill):
        pts = [cx, cy-h, cx+w, cy, cx, cy+h, cx-w, cy]
        canvas.create_polygon(pts, fill=fill, outline="", smooth=False)

    # Layout: 8 diamonds in the ETEC leaf pattern
    # Row 1 (top): teal pair
    diamond(x+40*s, y+12*s,  11*s, 10*s, ETEC_TEAL)
    diamond(x+58*s, y+12*s,  11*s, 10*s, ETEC_TEAL)
    # Row 2: green(left), blue(centre-left), blue(centre-right), purple(right)
    diamond(x+22*s, y+28*s,  11*s, 10*s, ETEC_GREEN)
    diamond(x+40*s, y+28*s,  11*s, 10*s, ETEC_BLUE)
    diamond(x+58*s, y+28*s,  11*s, 10*s, ETEC_BLUE)
    diamond(x+76*s, y+28*s,  11*s, 10*s, ETEC_PURPLE)
    # Row 3 (bottom): green(left), purple(right)
    diamond(x+31*s, y+44*s,  11*s, 10*s, ETEC_GREEN)
    diamond(x+67*s, y+44*s,  11*s, 10*s, ETEC_PURPLE)


class LauncherWindow(tk.Tk):
    """
    ETEC-branded launcher — full-screen dashboard with company identity.
    """
    def __init__(self):
        super().__init__()
        self.title("STGS — هيئة تقويم التعليم والتدريب")
        self.configure(bg=ETEC_NAVY)
        self.resizable(True, True)
        self.minsize(920, 620)
        self._build()

    # ── Build ──────────────────────────────────────────────────────────────
    def _build(self):
        self._build_header()
        self._build_tiles()
        self._build_footer()

    # ── Header ─────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=ETEC_NAVY); hdr.pack(fill="x")

        # Top accent line (teal)
        tk.Frame(hdr, bg=ETEC_TEAL, height=4).pack(fill="x")

        body = tk.Frame(hdr, bg=ETEC_NAVY, pady=18); body.pack(fill="x")

        # ── Logo canvas (left) ─────────────────────────────────────────────
        logo_c = tk.Canvas(body, bg=ETEC_NAVY, highlightthickness=0,
                           width=110, height=65)
        logo_c.pack(side="left", padx=(28, 0))
        _draw_etec_logo(logo_c, x=5, y=2, scale=1.05)

        # ── Text block (centre) ────────────────────────────────────────────
        txt = tk.Frame(body, bg=ETEC_NAVY); txt.pack(side="left", padx=18)

        tk.Label(txt, text="هيئة تقويم التعليم والتدريب",
                 bg=ETEC_NAVY, fg=ETEC_WHITE,
                 font=("Segoe UI", 20, "bold"), justify="left").pack(anchor="w")
        tk.Label(txt, text="Education & Training Evaluation Commission",
                 bg=ETEC_NAVY, fg=ETEC_TEAL,
                 font=("Segoe UI", 11), justify="left").pack(anchor="w")
        tk.Frame(txt, bg=ETEC_TEAL, height=2).pack(fill="x", pady=(6, 2))
        tk.Label(txt,
                 text="Smart Test Generation System  |  STGS",
                 bg=ETEC_NAVY, fg="#C8D8F0",
                 font=("Segoe UI", 10), justify="left").pack(anchor="w")
        tk.Label(txt,
                 text="إدارة الاختبارات الرقمية",
                 bg=ETEC_NAVY, fg=ETEC_PURPLE,
                 font=("Segoe UI", 10, "italic"), justify="left").pack(anchor="w")

        # (no version/creator tag in header — shown in footer only)

    # ── Tile area ───────────────────────────────────────────────────────────
    def _build_tiles(self):
        # Subtitle
        sub = tk.Frame(self, bg=ETEC_NAVY, pady=6); sub.pack(fill="x")
        tk.Label(sub, text="Select the exam system you want to use",
                 bg=ETEC_NAVY, fg="#8EA8CC",
                 font=("Segoe UI", 11)).pack()

        # Separator
        tk.Frame(self, bg=ETEC_BORDER, height=1).pack(fill="x", padx=30)

        outer = tk.Frame(self, bg=ETEC_LIGHT); outer.pack(fill="both", expand=True)

        grid = tk.Frame(outer, bg=ETEC_LIGHT); grid.pack(expand=True, pady=30)

        self._card(grid, col=0,
            accent=ETEC_BLUE,
            icon_text="📋",
            mode_tag="MODE 1",
            title="National & Standardised Exams",
            subtitle="القدرات · التحصيلي · القدرة المعرفية",
            bullets=[
                "قدرات علمي  ·  قدرات نظري",
                "التحصيلي  ·  القدرة المعرفية",
                "قدرات الجامعيين",
                "",
                "Column: Category  |  Domain: D",
                "Stages: Stage1 · E · M · D",
            ],
            command=self._open1)

        # Divider
        div = tk.Frame(grid, bg=ETEC_BORDER, width=1)
        div.grid(row=0, column=1, sticky="ns", padx=24, pady=10)

        self._card(grid, col=2,
            accent=ETEC_GREEN,
            icon_text="🏫",
            mode_tag="MODE 2",
            title="Educational & Curriculum Assessments",
            subtitle="المناهج · الصفوف الدراسية · المؤشرات",
            bullets=[
                "علوم  ·  رياضيات  ·  قراءة",
                "الصف الثالث  ·  السادس  ·  التاسع",
                "نافس  ·  القدرة المعرفية",
                "",
                "Outcome: الناتج  |  Indicator: المؤشر",
                "Filters: Subject · Grade · Language",
            ],
            command=self._open2)

        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(2, weight=1)
        grid.rowconfigure(0, weight=1)

    def _card(self, parent, col, accent, icon_text,
              mode_tag, title, subtitle, bullets, command):
        """Render one mode card with ETEC styling."""
        card = tk.Frame(parent, bg=ETEC_WHITE,
                        relief="flat", bd=0, cursor="hand2")
        card.grid(row=0, column=col, padx=18, sticky="nsew")

        # Coloured top bar
        tk.Frame(card, bg=accent, height=6).pack(fill="x")

        body = tk.Frame(card, bg=ETEC_WHITE, padx=28, pady=20)
        body.pack(fill="both", expand=True)

        # Icon + mode tag row
        top_row = tk.Frame(body, bg=ETEC_WHITE); top_row.pack(anchor="w")
        tk.Label(top_row, text=icon_text, bg=ETEC_WHITE,
                 font=("Segoe UI", 28)).pack(side="left", padx=(0, 10))
        tag_frame = tk.Frame(top_row, bg=accent, padx=8, pady=2)
        tag_frame.pack(side="left", anchor="s")
        tk.Label(tag_frame, text=mode_tag, bg=accent, fg=ETEC_WHITE,
                 font=("Segoe UI", 8, "bold")).pack()

        # Title
        tk.Label(body, text=title, bg=ETEC_WHITE, fg=ETEC_NAVY,
                 font=("Segoe UI", 13, "bold"),
                 wraplength=260, justify="left").pack(anchor="w", pady=(10, 2))

        # Arabic subtitle
        tk.Label(body, text=subtitle, bg=ETEC_WHITE, fg=accent,
                 font=("Segoe UI", 10, "italic")).pack(anchor="w", pady=(0, 10))

        # Separator
        tk.Frame(body, bg=ETEC_BORDER, height=1).pack(fill="x", pady=(0, 10))

        # Bullet points
        for line in bullets:
            if line == "":
                tk.Frame(body, bg=ETEC_WHITE, height=4).pack()
                continue
            row = tk.Frame(body, bg=ETEC_WHITE); row.pack(anchor="w", pady=1)
            tk.Label(row, text="▸", bg=ETEC_WHITE, fg=accent,
                     font=("Segoe UI", 9)).pack(side="left")
            tk.Label(row, text=f"  {line}", bg=ETEC_WHITE, fg="#444",
                     font=("Segoe UI", 9)).pack(side="left")

        # Open button
        btn = tk.Button(body, text="Open  →",
                        bg=accent, fg=ETEC_WHITE,
                        font=("Segoe UI", 10, "bold"),
                        relief="flat", padx=20, pady=7,
                        cursor="hand2", command=command,
                        activebackground=ETEC_NAVY,
                        activeforeground=ETEC_WHITE)
        btn.pack(anchor="w", pady=(16, 0))

        # Hover effect — lighten card border
        def _enter(_):
            card.config(highlightbackground=accent,
                        highlightthickness=2, relief="solid")
        def _leave(_):
            card.config(highlightthickness=0, relief="flat")
        card.bind("<Enter>", _enter); card.bind("<Leave>", _leave)
        card.bind("<Button-1>", lambda e: command())

    # ── Footer ──────────────────────────────────────────────────────────────
    def _build_footer(self):
        ftr = tk.Frame(self, bg=ETEC_NAVY, pady=6)
        ftr.pack(fill="x", side="bottom")

        # Teal bottom accent line
        tk.Frame(ftr, bg=ETEC_TEAL, height=3).pack(fill="x", side="bottom")

        # Features line
        tk.Label(ftr,
                 text="Stratified Bell Sampling  ·  Mean-Criterion Retry  ·  "
                      "3PL Auto-Analysis  ·  Professional Excel Output",
                 bg=ETEC_NAVY, fg="#6A86AA",
                 font=("Segoe UI", 8)).pack()

        # Department + company + creator — all on one row
        row = tk.Frame(ftr, bg=ETEC_NAVY); row.pack(pady=(2, 0))

        tk.Label(row,
                 text="إدارة الاختبارات الرقمية  |  ",
                 bg=ETEC_NAVY, fg=ETEC_PURPLE,
                 font=("Segoe UI", 8, "italic")).pack(side="left")

        tk.Label(row,
                 text="© هيئة تقويم التعليم والتدريب  —  Education & Training Evaluation Commission",
                 bg=ETEC_NAVY, fg="#4A5C78",
                 font=("Segoe UI", 8)).pack(side="left")

        tk.Label(row,
                 text="  |  Created by  ",
                 bg=ETEC_NAVY, fg="#4A5C78",
                 font=("Segoe UI", 8)).pack(side="left")

        tk.Label(row,
                 text="Saeed Alkaltham",
                 bg=ETEC_NAVY, fg=ETEC_TEAL,
                 font=("Segoe UI", 8, "bold")).pack(side="left")

    # ── Navigation ──────────────────────────────────────────────────────────
    def _open1(self):
        self.withdraw()
        w = Mode1Window(on_close=self.deiconify)
        w.protocol("WM_DELETE_WINDOW", lambda: (w.destroy(), self.deiconify()))

    def _open2(self):
        self.withdraw()
        w = Mode2Window(on_close=self.deiconify)
        w.protocol("WM_DELETE_WINDOW", lambda: (w.destroy(), self.deiconify()))

# ─────────────────────────────────────────────────────────────────────────────
#  MODE 1 WINDOW
# ─────────────────────────────────────────────────────────────────────────────

class _BaseMode(tk.Toplevel):
    """Shared behaviour for both mode windows."""
    COLOR = BG_M

    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait(); k = item[0]
                if k == "bank":
                    self._on_bank(item[1], item[2])
                elif k == "log":
                    _append_log(self._logbox, item[1], item[2] if len(item) > 2 else "")
                elif k == "prog":
                    self._pv.set(item[1])
                elif k == "done":
                    self.analyses = item[1]
                    self._gbtn.config(state="normal"); self._pv.set(100)
                    self._refresh_results()
                    self._nb.select(self._results_tab_idx)
                    self._sv.set("Generation complete.")
                    mb.showinfo("Done",
                                "All forms generated and analysed.\n"
                                "Output files saved to the same folder as the bank.")
                elif k == "fail":
                    self._gbtn.config(state="normal")
                    self._sv.set("Generation failed — see log.")
                elif k == "err":
                    self._sv.set(f"Error: {item[1]}")
                    mb.showerror("Error", item[1])
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _close(self):
        self.destroy()
        if self._on_close: self._on_close()

    def _open_folder(self):
        if self.source_path:
            d = str(self.source_path.parent)
            if sys.platform == "win32": subprocess.Popen(["explorer", d])
            elif sys.platform == "darwin": subprocess.Popen(["open", d])
            else: subprocess.Popen(["xdg-open", d])

    def _le(self, p, lbl, row, default=""):
        tk.Label(p, text=lbl, bg=BG_L, fg=TXD, font=FB, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        e = _entry(p, w=16); e.insert(0, default)
        e.grid(row=row, column=1, sticky="w", pady=2)
        return e

    def _lg(self, msg, col=""): self._q.put(("log", msg, col))


class Mode1Window(_BaseMode):
    COLOR = ETEC_BLUE
    _results_tab_idx = 3

    def __init__(self, on_close=None):
        super().__init__()
        self.title("STGS — Mode 1 | هيئة تقويم التعليم والتدريب")
        self.configure(bg=ETEC_NAVY); self.resizable(True, True); self.minsize(980, 700)
        self._on_close = on_close
        self.bank_df: Optional[pd.DataFrame] = None
        self.source_path: Optional[Path] = None
        self.analyses: list[FormAnalysis] = []
        self._q: queue.Queue = queue.Queue()
        self._sub_rows:  list[tuple[str, tk.Entry]] = []
        self._bins_map:  dict[str, list] = {}
        self._build(); self.after(80, self._poll)

    # ── build ─────────────────────────────────────────────────────────────────
    def _build(self):
        # ── Top header bar ────────────────────────────────────────────────────
        hf = tk.Frame(self, bg=ETEC_NAVY); hf.pack(fill="x")
        tk.Frame(hf, bg=ETEC_TEAL, height=4).pack(fill="x")   # teal accent line
        hf2 = tk.Frame(hf, bg=ETEC_NAVY, pady=8); hf2.pack(fill="x")

        # Mini logo canvas
        lc = tk.Canvas(hf2, bg=ETEC_NAVY, highlightthickness=0, width=64, height=44)
        lc.pack(side="left", padx=(12, 0))
        _draw_etec_logo(lc, x=2, y=1, scale=0.65)

        txt = tk.Frame(hf2, bg=ETEC_NAVY); txt.pack(side="left", padx=10)
        tk.Label(txt, text="Mode 1 — National & Standardised Exams",
                 bg=ETEC_NAVY, fg=ETEC_WHITE, font=FT).pack(anchor="w")
        tk.Label(txt, text="هيئة تقويم التعليم والتدريب  |  STGS",
                 bg=ETEC_NAVY, fg=ETEC_TEAL, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(txt,
                 text="إدارة الاختبارات الرقمية",
                 bg=ETEC_NAVY, fg=ETEC_PURPLE,
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")

        _btn(hf2, "← Back", self._close, bg=ETEC_PURPLE).pack(
            side="right", padx=12, pady=4)

        # ── Notebook ──────────────────────────────────────────────────────────
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("M1.TNotebook", background=ETEC_NAVY, borderwidth=0)
        s.configure("M1.TNotebook.Tab",
                    background=ETEC_BLUE, foreground=TXL, font=FH, padding=[12, 5])
        s.map("M1.TNotebook.Tab",
              background=[("selected", ETEC_TEAL)],
              foreground=[("selected", ETEC_WHITE)])
        self._nb = ttk.Notebook(self, style="M1.TNotebook")
        self._nb.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        t1=_scroll_frame(self._nb); t2=_scroll_frame(self._nb)
        t3=tk.Frame(self._nb, bg=BG_L); t4=_scroll_frame(self._nb)
        self._nb.add(t1, text="  📂 Bank  "); self._nb.add(t2, text="  ⚙ Configure  ")
        self._nb.add(t3, text="  ▶ Generate  "); self._nb.add(t4, text="  📊 Results  ")

        self._build_bank_tab(t1._inner)
        self._build_config_tab(t2._inner)
        self._build_generate_tab(t3)
        self._build_results_tab(t4._inner)

        # ── Status bar ────────────────────────────────────────────────────────
        sf = tk.Frame(self, bg=ETEC_NAVY, height=22); sf.pack(fill="x", side="bottom")
        tk.Frame(sf, bg=ETEC_TEAL, height=2).pack(fill="x", side="bottom")
        self._sv = tk.StringVar(value="Ready")
        tk.Label(sf, textvariable=self._sv, bg=ETEC_NAVY, fg=ETEC_TEAL,
                 font=FM, anchor="w", padx=8).pack(fill="x")

    # ── Bank tab ──────────────────────────────────────────────────────────────
    def _build_bank_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "Question Bank", bold=True).pack(anchor="w", pady=(0, 6))
        uf = tk.Frame(p, bg=BG_L); uf.pack(fill="x", pady=4)
        self._bp = tk.StringVar(value="No file selected")
        tk.Label(uf, textvariable=self._bp, bg=BG_L, fg=BG_D,
                 font=FB, width=55, anchor="w").pack(side="left")
        _btn(uf, "Upload Excel / CSV", self._browse_bank,
             bg=self.COLOR).pack(side="left", padx=8)
        self._bstats = tk.StringVar(value="")
        tk.Label(p, textvariable=self._bstats, bg=BG_L, fg=BG_M,
                 font=FB, justify="left").pack(anchor="w", pady=4)
        _lbl(p, "Categories detected in 'Category' column", bold=True).pack(
            anchor="w", pady=(10, 4))
        self._cat_tree = _treeview(p,
            ["Category","Domain (D)","Count","Mean Difficulty","Min","Max"],
            [160, 140, 60, 120, 65, 65])

    def _browse_bank(self):
        path = fd.askopenfilename(title="Select Question Bank",
            filetypes=[("Excel / CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not path: return
        self._sv.set("Loading…")
        threading.Thread(target=lambda: self._load_t(path), daemon=True).start()

    def _load_t(self, path):
        try:
            df = load_bank_mode1(path); self._q.put(("bank", df, path))
        except Exception as e: self._q.put(("err", str(e)))

    def _on_bank(self, df, path):
        self.bank_df = df; self.source_path = Path(path); self._bp.set(Path(path).name)
        self._bstats.set(
            f"Questions: {len(df)}   |   Categories: {df['Category'].nunique()}   |   "
            f"Difficulty: [{df['Difficulty'].min():.2f}, {df['Difficulty'].max():.2f}]   |   "
            f"Mean: {df['Difficulty'].mean():.3f}")
        for i in self._cat_tree.get_children(): self._cat_tree.delete(i)
        for cat, g in df.groupby("Category"):
            dom = g["D"].mode().iloc[0] if "D" in g else ""
            self._cat_tree.insert("","end", values=[
                cat, dom, len(g),
                f"{g['Difficulty'].mean():.3f}",
                f"{g['Difficulty'].min():.3f}",
                f"{g['Difficulty'].max():.3f}"])
        self._populate_subdomain_entries()
        self._sv.set(f"Loaded: {len(df)} questions")

    # ── Config tab ────────────────────────────────────────────────────────────
    def _build_config_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        p.columnconfigure(0, weight=1); p.columnconfigure(1, weight=1)

        # Exam template
        ef = tk.LabelFrame(p, text="Exam Template",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        ef.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=4)
        r0 = tk.Frame(ef, bg=BG_L); r0.pack(fill="x")
        tk.Label(r0, text="Exam type:", bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._exam_var = tk.StringVar(value=list(MODE1_EXAMS.keys())[0])
        cb = ttk.Combobox(r0, textvariable=self._exam_var,
                          values=list(MODE1_EXAMS.keys()), state="readonly", width=28)
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda e: self._populate_subdomain_entries())
        _btn(r0, "Apply", self._populate_subdomain_entries, bg=BG_D).pack(side="left", padx=4)

        # Subdomain header
        hdr = tk.Frame(ef, bg=BG_L); hdr.pack(fill="x", pady=(6, 2))
        for i, (t, w) in enumerate([("Category", 22), ("Count", 9)]):
            tk.Label(hdr, text=t, bg=BG_M, fg=TXL, font=FH,
                     width=w, relief="flat", padx=4).grid(row=0, column=i, padx=2)
        self._sub_frame = tk.Frame(ef, bg=BG_L); self._sub_frame.pack(fill="x")

        # Assembly params
        pf = tk.LabelFrame(p, text="Assembly Parameters",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        pf.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=4)
        self._nforms = self._le(pf, "Number of Forms:", 0, "4")

        tk.Label(pf, text="Difficulty Stage:", bg=BG_L, fg=TXD, font=FB).grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=2)
        self._stage_var = tk.StringVar(value=list(MODE1_STAGES.keys())[0])
        self._stage_var.trace_add("write", self._on_stage)
        ttk.Combobox(pf, textvariable=self._stage_var,
                     values=list(MODE1_STAGES.keys()),
                     state="readonly", width=22).grid(row=1, column=1, sticky="w", pady=2)

        self._rng_lbl = tk.Label(pf, text="", bg=BG_L, fg=TXD, font=FB)
        self._rng_ent = _entry(pf, w=20); self._rng_ent.config(state="disabled")
        self._mean_lbl = tk.Label(pf, text="", bg=BG_L, fg=TXD, font=FB)
        self._mean_ent = _entry(pf, w=20); self._mean_ent.config(state="disabled")
        self._rng_lbl.grid(row=2, column=0, sticky="w"); self._rng_ent.grid(row=2, column=1, sticky="w")
        self._mean_lbl.grid(row=3, column=0, sticky="w"); self._mean_ent.grid(row=3, column=1, sticky="w")

        self._stats_var   = tk.BooleanVar(value=False)
        self._partial_var = tk.BooleanVar(value=True)
        self._reuse_var   = tk.BooleanVar(value=False)
        for row, (var, text) in enumerate([
            (self._stats_var,   "Enable discrimination filter (Discrimination ≥ 0.85)"),
            (self._partial_var, "Allow partial fill (placeholder rows when short)"),
        ], start=4):
            tk.Checkbutton(pf, text=text, variable=var,
                           bg=BG_L, fg=TXD, font=FB).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=2)

        # Reuse row: checkbox + "Fill in:" label + spinbox on the same line
        reuse_frm = tk.Frame(pf, bg=BG_L)
        reuse_frm.grid(row=6, column=0, columnspan=2, sticky="w", pady=2)
        tk.Checkbutton(reuse_frm, text="Allow reuse across forms — never in same form",
                       variable=self._reuse_var,
                       bg=BG_L, fg=TXD, font=FB).pack(side="left")
        tk.Label(reuse_frm, text="   Fill in:",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._max_reuse_m1 = tk.Spinbox(
            reuse_frm, from_=2, to=10, width=4, font=FB,
            bg="#FFF", fg="#000", relief="solid", bd=1)
        self._max_reuse_m1.delete(0, "end"); self._max_reuse_m1.insert(0, "2")
        self._max_reuse_m1.pack(side="left", padx=4)
        tk.Label(reuse_frm, text="times max per question",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")

        self._sim_n = self._le(pf, "Simulation students:", 7, "3000")
        self._seed  = self._le(pf, "RNG Seed (blank=random):", 8, "")

        # Save settings button
        _btn(pf, "Save & Preview Settings", self._save_settings,
             bg=BG_D).grid(row=9, column=0, columnspan=2, pady=8)

        # Manual bins
        bf = tk.LabelFrame(p, text="Manual Difficulty Bins (0.0 → 1.0)  — optional",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        bf.grid(row=1, column=0, columnspan=2, sticky="ew", pady=4)
        self._use_bins = tk.BooleanVar(value=False)
        tk.Checkbutton(bf, text="Enable manual bin distribution",
                       variable=self._use_bins, bg=BG_L, fg=TXD, font=FB).pack(anchor="w")
        self._bins_container = tk.Frame(bf, bg=BG_L); self._bins_container.pack(fill="x")

    def _populate_subdomain_entries(self, *_):
        for w in self._sub_frame.winfo_children(): w.destroy()
        self._sub_rows.clear()
        exam = self._exam_var.get()
        for cat, cnt in MODE1_EXAMS.get(exam, {}).items():
            frm = tk.Frame(self._sub_frame, bg=BG_L); frm.pack(fill="x", pady=1)
            tk.Label(frm, text=cat, bg=BG_L, fg=TXD, font=FB,
                     width=28, anchor="w").grid(row=0, column=0, padx=4)
            e = _entry(frm, w=8); e.insert(0, str(cnt))
            e.grid(row=0, column=1, padx=4)
            self._sub_rows.append((cat, e))
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
                    tk.Label(frm, text=BIN_LABELS[idx]+":", bg=BG_L,
                             fg=TXD, font=FM, width=9).grid(row=0, column=bi*2, padx=2)
                    e = _entry(frm, w=4); e.grid(row=0, column=bi*2+1, padx=2)
                    entries.append(e)
            self._bins_map[cat] = entries

    def _on_stage(self, *_):
        is_m = "Manually" in self._stage_var.get()
        if is_m:
            self._rng_lbl.config(text="Range (e.g. 0.2, 0.8):")
            self._rng_ent.config(state="normal")
            self._mean_lbl.config(text="Mean range (e.g. 0.5, 0.6):")
            self._mean_ent.config(state="normal")
        else:
            self._rng_lbl.config(text=""); self._rng_ent.config(state="disabled")
            self._mean_lbl.config(text=""); self._mean_ent.config(state="disabled")

    def _save_settings(self):
        """Preview / confirm current settings (mirrors save_settings() from Script 1)."""
        try:
            n_forms = int(self._nforms.get())
        except ValueError:
            n_forms = "?"
        stage = self._stage_var.get()
        crit  = MODE1_STAGES.get(stage, {})
        exam  = self._exam_var.get()
        total = sum(int(e.get()) for _, e in self._sub_rows
                    if e.get().isdigit())
        lines = [
            f"Exam        : {exam}",
            f"Stage       : {stage}",
            f"Diff Range  : {crit.get('Range', 'manual')}",
            f"Mean Target : {crit.get('Mean', 'manual')}",
            f"Forms       : {n_forms}",
            f"Total Q/form: {total}",
            f"Bins        : {'Enabled' if self._use_bins.get() else 'Disabled'}",
            f"Disc filter : {'Enabled (≥0.85)' if self._stats_var.get() else 'Disabled'}",
            f"Partial fill: {'Enabled' if self._partial_var.get() else 'Disabled'}",
            f"Reuse       : {'Enabled — max ' + self._max_reuse_m1.get() + 'x per question' if self._reuse_var.get() else 'Disabled'}",
        ]
        mb.showinfo("Settings Preview", "\n".join(lines))

    def _validate_manual_bins(self) -> bool:
        """Validate bin totals match subdomain counts (mirrors validate_manual_distribution)."""
        if not self._use_bins.get(): return True
        errors = []
        for cat, ent in self._sub_rows:
            try: expected = int(ent.get())
            except: expected = 0
            if cat not in self._bins_map: continue
            total_bins = sum(
                int(e.get()) for e in self._bins_map[cat] if e.get().isdigit()
            )
            if total_bins > 0 and total_bins != expected:
                errors.append(f"  {cat}: bins sum to {total_bins}, expected {expected}")
        if errors:
            mb.showerror("Bin Validation Failed",
                         "The following categories have mismatched bin totals:\n" +
                         "\n".join(errors) +
                         "\n\nCorrect the bin counts or disable manual bins.")
            return False
        return True

    # ── Generate tab ──────────────────────────────────────────────────────────
    def _build_generate_tab(self, p):
        p.configure(bg=BG_L)
        tk.Label(p, text="Generate Forms & Run 3PL Analysis",
                 bg=BG_L, fg=TXD, font=FH).pack(anchor="w", padx=20, pady=(14, 4))
        bf = tk.Frame(p, bg=BG_L, pady=6); bf.pack(padx=20, anchor="w")
        self._gbtn = _btn(bf, "▶  Generate & Analyse", self._run, bg=self.COLOR)
        self._gbtn.config(font=("Segoe UI", 11, "bold"), pady=8, padx=20)
        self._gbtn.pack(side="left", padx=(0, 12))
        self._pv = tk.DoubleVar(value=0)
        ttk.Progressbar(bf, variable=self._pv, maximum=100,
                        length=320, mode="determinate").pack(side="left")
        lf = tk.LabelFrame(p, text="Log", bg=BG_L, fg=BG_D,
                           font=FH, padx=8, pady=8)
        lf.pack(fill="both", expand=True, padx=20, pady=8)
        self._logbox = _log_widget(lf)

    # ── Results tab ───────────────────────────────────────────────────────────
    def _build_results_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "Results Summary", bold=True).pack(anchor="w", pady=(0, 8))
        self._rtree = _treeview(p,
            ["Form","Items","Mean Diff","Min","Max","Avg a","Avg c","Alpha","Item Corr."],
            [90, 60, 100, 60, 60, 75, 65, 100, 100])
        bf = tk.Frame(p, bg=BG_L); bf.pack(anchor="w", pady=8)
        _btn(bf, "Open Output Folder", self._open_folder, bg=self.COLOR).pack(side="left")
        _btn(bf, "Refresh", self._refresh_results, bg=BG_D).pack(side="left", padx=6)

    def _refresh_results(self):
        for i in self._rtree.get_children(): self._rtree.delete(i)
        for fa in self.analyses:
            self._rtree.insert("", "end", values=[
                fa.name, fa.n_items,
                f"{fa.mean_diff:.3f}", f"{fa.min_diff:.3f}", f"{fa.max_diff:.3f}",
                f"{fa.avg_a:.3f}", f"{fa.avg_c:.3f}",
                f"{fa.alpha:.3f}" if not math.isnan(fa.alpha) else "—",
                f"{fa.item_corr:.3f}" if not math.isnan(fa.item_corr) else "—",
            ])

    # ── Run logic ─────────────────────────────────────────────────────────────
    def _run(self):
        if self.bank_df is None:
            mb.showerror("Error", "Please upload a question bank first (Bank tab)."); return
        if not self._validate_manual_bins(): return
        try: params = self._collect()
        except ValueError as e: mb.showerror("Configuration Error", str(e)); return
        self._gbtn.config(state="disabled"); self._pv.set(0)
        self._logbox.config(state="normal"); self._logbox.delete("1.0","end")
        self._logbox.config(state="disabled")
        threading.Thread(target=self._gen_t, args=(params,), daemon=True).start()

    def _collect(self) -> dict:
        def ii(e, n):
            try: return int(e.get().strip())
            except: raise ValueError(f"Invalid value for '{n}'")

        n_forms = ii(self._nforms, "Number of Forms")
        stage   = self._stage_var.get()
        crit    = MODE1_STAGES.get(stage, {})
        diff_range = crit.get("Range") or (0.0, 1.0)
        mean_range = crit.get("Mean")
        if "Manually" in stage:
            try:
                parts = [float(x) for x in self._rng_ent.get().split(",")]
                if len(parts) == 2: diff_range = tuple(parts)
            except: pass
            try:
                parts = [float(x) for x in self._mean_ent.get().split(",")]
                if len(parts) == 2: mean_range = tuple(parts)
            except: pass

        slots = []
        for cat, ent in self._sub_rows:
            try: cnt = int(ent.get())
            except: cnt = 0
            if cnt <= 0: continue
            bins = []
            if self._use_bins.get() and cat in self._bins_map:
                for bi, e in enumerate(self._bins_map[cat]):
                    try: bc = int(e.get())
                    except: bc = 0
                    if bc > 0:
                        bins.append(BinDef(bi*0.1, (bi+1)*0.1, bc))
            slots.append(SlotDef(label=cat,
                                  filters={"Category": cat},
                                  count=cnt, bins=bins))
        if not slots:
            raise ValueError("No categories configured. Select an exam template.")

        seed_s = self._seed.get().strip()
        try:
            max_reuse_val = max(2, int(self._max_reuse_m1.get()))
        except ValueError:
            max_reuse_val = 2
        params = AssemblyParams(
            slots=slots,
            diff_range=diff_range,
            diff_mean_range=mean_range,
            min_discrimination=0.85 if self._stats_var.get() else None,
            allow_partial_fill=self._partial_var.get(),
            allow_reuse=self._reuse_var.get(),
            max_reuse=max_reuse_val,
            max_retries=100,
            rng_seed=int(seed_s) if seed_s else None,
        )
        return {"n_forms": n_forms, "params": params,
                "n_students": ii(self._sim_n, "Simulation students"),
                "seed": int(seed_s) if seed_s else None}

    def _gen_t(self, p):
        try:
            n = p["n_forms"]; params = p["params"]
            names = [f"Form_{i+1}" for i in range(n)]
            self._lg(f"Starting assembly: {n} forms…")
            results = assemble_forms(self.bank_df, params, "Difficulty", n, names)
            self._q.put(("prog", 40))
            for (form, mean_d, warns), name in zip(results, names):
                for w in warns: self._lg(f"  ⚠ {w}", "yellow")
                self._lg(f"  {name}: {len(form)} items, mean difficulty = {mean_d:.3f}")

            self._lg("Running 3PL psychometric analysis…")
            analyses = [
                analyse_form(r[0], nm, "Difficulty", "Discrimination", "Guessing",
                             p["n_students"], p["seed"])
                for r, nm in zip(results, names)
            ]
            self._q.put(("prog", 75))
            for fa in analyses:
                if fa:
                    self._lg(f"  {fa.name}: Alpha = {fa.alpha:.3f}, "
                             f"Item Corr. = {fa.item_corr:.3f}")

            self._lg("Writing output files…")
            base = self.source_path.parent; prefix = self.source_path.stem
            fp = base / f"{prefix}_Forms.xlsx"
            ap = base / f"{prefix}_3PL_Analysis.xlsx"
            rp = base / f"{prefix}_Remaining_Questions.xlsx"

            # Usage tracking: by QuestionID
            usage: dict[str, int] = {}
            for form, _, _ in results:
                if "QuestionID" in form.columns:
                    for qid in form["QuestionID"].dropna().astype(str):
                        usage[qid] = usage.get(qid, 0) + 1

            write_forms_workbook(
                results, names, analyses,
                "Difficulty", "D", "Discrimination", "Guessing",
                usage, fp, mode=1)
            valid = [fa for fa in analyses if fa]
            if valid: write_analysis_workbook(valid, ap)

            used_ids = set()
            for form, _, _ in results:
                if "QuestionID" in form.columns:
                    used_ids.update(form["QuestionID"].dropna().astype(str).tolist())
            remaining = self.bank_df[
                ~self.bank_df["QuestionID"].astype(str).isin(used_ids)
            ].copy()
            remaining.to_excel(str(rp), index=False, engine="openpyxl")

            self._q.put(("prog", 100))
            self._lg(f"\n  📄 Forms           : {fp}")
            self._lg(f"  📊 3PL Analysis    : {ap}")
            self._lg(f"  📋 Remaining Qs    : {rp}")
            self._lg("\n✓ Generation complete!", "lime")
            self._q.put(("done", valid))
        except Exception as e:
            self._lg(f"\n✗ Error: {e}", "red")
            log.exception("Mode1 generation failed")
            self._q.put(("fail",))

# ─────────────────────────────────────────────────────────────────────────────
#  MODE 2 WINDOW
# ─────────────────────────────────────────────────────────────────────────────

class Mode2Window(_BaseMode):
    COLOR = ETEC_GREEN
    _results_tab_idx = 4

    def __init__(self, on_close=None):
        super().__init__()
        self.title("STGS — Mode 2 | هيئة تقويم التعليم والتدريب")
        self.configure(bg=ETEC_NAVY); self.resizable(True, True); self.minsize(980, 700)
        self._on_close   = on_close
        self.bank_df:      Optional[pd.DataFrame] = None
        self._filtered_df: Optional[pd.DataFrame] = None
        self.source_path:  Optional[Path] = None
        self.analyses:     list[FormAnalysis] = []
        self._q:           queue.Queue = queue.Queue()
        self._indicator_entries: list[tuple[str, dict[str, tk.Entry]]] = []
        self._bins_map:   dict[str, list] = {}
        self._build(); self.after(80, self._poll)

    def _build(self):
        # ── Top header bar ────────────────────────────────────────────────────
        hf = tk.Frame(self, bg=ETEC_NAVY); hf.pack(fill="x")
        tk.Frame(hf, bg=ETEC_GREEN, height=4).pack(fill="x")   # green accent line
        hf2 = tk.Frame(hf, bg=ETEC_NAVY, pady=8); hf2.pack(fill="x")

        lc = tk.Canvas(hf2, bg=ETEC_NAVY, highlightthickness=0, width=64, height=44)
        lc.pack(side="left", padx=(12, 0))
        _draw_etec_logo(lc, x=2, y=1, scale=0.65)

        txt = tk.Frame(hf2, bg=ETEC_NAVY); txt.pack(side="left", padx=10)
        tk.Label(txt, text="Mode 2 — Educational & Curriculum Assessments",
                 bg=ETEC_NAVY, fg=ETEC_WHITE, font=FT).pack(anchor="w")
        tk.Label(txt, text="هيئة تقويم التعليم والتدريب  |  STGS",
                 bg=ETEC_NAVY, fg=ETEC_GREEN, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(txt,
                 text="إدارة الاختبارات الرقمية",
                 bg=ETEC_NAVY, fg=ETEC_PURPLE,
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")

        _btn(hf2, "← Back", self._close, bg=ETEC_PURPLE).pack(
            side="right", padx=12, pady=4)

        # ── Notebook ──────────────────────────────────────────────────────────
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("M2.TNotebook", background=ETEC_NAVY, borderwidth=0)
        s.configure("M2.TNotebook.Tab",
                    background=ETEC_GREEN, foreground=TXL, font=FH, padding=[12, 5])
        s.map("M2.TNotebook.Tab",
              background=[("selected", ETEC_TEAL)],
              foreground=[("selected", ETEC_WHITE)])
        self._nb = ttk.Notebook(self, style="M2.TNotebook")
        self._nb.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        t1=_scroll_frame(self._nb); t2=_scroll_frame(self._nb)
        t3=_scroll_frame(self._nb); t4=tk.Frame(self._nb, bg=BG_L)
        t5=_scroll_frame(self._nb)
        self._nb.add(t1, text="  📂 Bank  ");  self._nb.add(t2, text="  ⚙ Configure  ")
        self._nb.add(t3, text="  📐 Bins  ");  self._nb.add(t4, text="  ▶ Generate  ")
        self._nb.add(t5, text="  📊 Results  ")

        self._build_bank_tab(t1._inner)
        self._build_config_tab(t2._inner)
        self._build_bins_tab(t3._inner)
        self._build_generate_tab(t4)
        self._build_results_tab(t5._inner)

        # ── Status bar ────────────────────────────────────────────────────────
        sf = tk.Frame(self, bg=ETEC_NAVY, height=22); sf.pack(fill="x", side="bottom")
        tk.Frame(sf, bg=ETEC_GREEN, height=2).pack(fill="x", side="bottom")
        self._sv = tk.StringVar(value="Ready")
        tk.Label(sf, textvariable=self._sv, bg=ETEC_NAVY, fg=ETEC_GREEN,
                 font=FM, anchor="w", padx=8).pack(fill="x")

    # ── Bank tab ──────────────────────────────────────────────────────────────
    def _build_bank_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "Question Bank", bold=True).pack(anchor="w", pady=(0, 6))
        uf = tk.Frame(p, bg=BG_L); uf.pack(fill="x", pady=4)
        self._bp = tk.StringVar(value="No file selected")
        tk.Label(uf, textvariable=self._bp, bg=BG_L, fg=BG_D,
                 font=FB, width=55, anchor="w").pack(side="left")
        _btn(uf, "Upload Excel / CSV", self._browse_bank,
             bg=self.COLOR).pack(side="left", padx=8)
        self._bstats = tk.StringVar(value="")
        tk.Label(p, textvariable=self._bstats, bg=BG_L, fg=BG_M,
                 font=FB, justify="left").pack(anchor="w", pady=4)

        # Filters — shown only when detected
        self._ffrm = tk.LabelFrame(p, text="Filter Bank",
                                   bg=BG_L, fg=BG_D, font=FH, padx=8, pady=6)
        self._ffrm.pack(fill="x", pady=4); self._ffrm.pack_forget()
        self._fsub = tk.StringVar(value="All")
        self._fgrd = tk.StringVar(value="All")
        self._flng = tk.StringVar(value="All")
        for i, (lbl, var, attr) in enumerate([
            ("Subject:", self._fsub, "_cb_sub"),
            ("Grade:",   self._fgrd, "_cb_grd"),
            ("Language:",self._flng, "_cb_lng"),
        ]):
            tk.Label(self._ffrm, text=lbl, bg=BG_L, fg=TXD,
                     font=FB).grid(row=0, column=i*2, padx=(6,2))
            cb = ttk.Combobox(self._ffrm, textvariable=var,
                              state="readonly", width=14)
            cb.grid(row=0, column=i*2+1, padx=(0, 10))
            cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())
            setattr(self, attr, cb)

        _lbl(p, "Outcome / Indicator Summary", bold=True).pack(anchor="w", pady=(10, 4))
        self._ptree = _treeview(p,
            ["Outcome (الناتج)","Indicator (المؤشر)","Count","Mean Diff","Min","Max"],
            [180, 180, 60, 110, 65, 65])

    def _browse_bank(self):
        path = fd.askopenfilename(title="Select Question Bank",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not path: return
        self._sv.set("Loading…")
        threading.Thread(target=lambda: self._load_t(path), daemon=True).start()

    def _load_t(self, path):
        try:
            df = load_bank_mode2(path); self._q.put(("bank", df, path))
        except Exception as e: self._q.put(("err", str(e)))

    def _on_bank(self, df, path):
        self.bank_df = df; self._filtered_df = df.copy()
        self.source_path = Path(path); self._bp.set(Path(path).name)
        has_f = all(c in df.columns for c in ("subject","grade","language"))
        if has_f:
            self._ffrm.pack(fill="x", pady=4)
            for col, cb in [("subject",self._cb_sub),("grade",self._cb_grd),
                            ("language",self._cb_lng)]:
                cb["values"] = ["All"] + sorted(df[col].dropna().astype(str).unique())
        else:
            self._ffrm.pack_forget()
        self._refresh_from_df(df)
        self._sv.set(f"Loaded: {len(df)} questions")

    def _refresh_from_df(self, df):
        self._bstats.set(
            f"Questions: {len(df)}   |   Pairs: {df['_pair_key'].nunique()}   |   "
            f"Difficulty: [{df['difficulty'].min():.2f}, {df['difficulty'].max():.2f}]   |   "
            f"Mean: {df['difficulty'].mean():.3f}")
        self._populate_pair_tree(df)
        self._update_subcategories(df)   # replicates update_subcategories()
        self._populate_bins(df)

    def _populate_pair_tree(self, df):
        for i in self._ptree.get_children(): self._ptree.delete(i)
        grp = df.groupby("_pair_key", sort=False)
        summ = grp.agg(
            الناتج=("الناتج","first"), المؤشر=("المؤشر","first"),
            count=("QuestionID","count"),
            mean_d=("difficulty","mean"),
            min_d=("difficulty","min"),
            max_d=("difficulty","max"),
        ).reset_index(drop=True)
        for _, row in summ.iterrows():
            self._ptree.insert("","end",values=[
                row["الناتج"], row["المؤشر"], row["count"],
                f"{row['mean_d']:.3f}", f"{row['min_d']:.3f}", f"{row['max_d']:.3f}"])

    def _apply_filters(self):
        if self.bank_df is None: return
        df = self.bank_df.copy()
        for col, var in [("subject",self._fsub),("grade",self._fgrd),("language",self._flng)]:
            if col in df.columns and var.get() != "All":
                df = df[df[col].astype(str) == var.get()]
        self._filtered_df = df
        self._refresh_from_df(df)

    # ── update_subcategories() from Script 2 ──────────────────────────────────
    def _update_subcategories(self, df):
        """
        Dynamically show (الناتج, المؤشر) pairs from the CURRENT filtered bank.
        If any indicator count is filled → pair-mode assembly.
        Otherwise → outcome-only mode.
        This directly replicates update_subcategories() from Script 2.
        """
        for w in self._ind_frame.winfo_children(): w.destroy()
        self._indicator_entries.clear()

        if "الناتج" not in df.columns or "المؤشر" not in df.columns:
            return

        for outcome in df["الناتج"].dropna().unique():
            hdr = tk.Label(self._ind_frame,
                           text=f"Outcome: {outcome}",
                           bg=BG_L, fg=BG_D, font=FH)
            hdr.pack(anchor="w", padx=8, pady=(8,2))
            indicators = df[df["الناتج"] == outcome]["المؤشر"].dropna().unique()
            ind_dict: dict[str, tk.Entry] = {}
            for ind in indicators:
                frm = tk.Frame(self._ind_frame, bg=BG_L); frm.pack(fill="x", padx=20, pady=1)
                tk.Label(frm, text=f"  Indicator {ind}:", bg=BG_L,
                         fg=TXD, font=FB, width=24, anchor="w").pack(side="left")
                e = _entry(frm, w=8); e.pack(side="left", padx=4)
                ind_dict[str(ind)] = e
            self._indicator_entries.append((str(outcome), ind_dict))

    # ── Config tab ────────────────────────────────────────────────────────────
    def _build_config_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        p.columnconfigure(0, weight=1); p.columnconfigure(1, weight=1)

        # Exam template
        ef = tk.LabelFrame(p, text="Exam Template (used to pre-fill outcome counts)",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        ef.grid(row=0, column=0, sticky="nsew", padx=(0,8), pady=4)
        r0 = tk.Frame(ef, bg=BG_L); r0.pack(fill="x")
        tk.Label(r0, text="Exam:", bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._exam_var = tk.StringVar(value=list(MODE2_EXAMS.keys())[0])
        cb = ttk.Combobox(r0, textvariable=self._exam_var,
                          values=list(MODE2_EXAMS.keys()), state="readonly", width=26)
        cb.pack(side="left", padx=6)

        # Outcome (الناتج) distribution header
        _lbl(ef, "Learning Outcome Distribution", bold=True).pack(anchor="w", pady=(8,2))
        hdr = tk.Frame(ef, bg=BG_L); hdr.pack(fill="x", pady=(0,2))
        for i,(t,w) in enumerate([("Learning Outcome (الناتج)",24),("Count",9)]):
            tk.Label(hdr, text=t, bg=BG_M, fg=TXL, font=FH,
                     width=w, relief="flat", padx=4).grid(row=0, column=i, padx=2)
        self._out_frame = tk.Frame(ef, bg=BG_L); self._out_frame.pack(fill="x")
        self._out_rows: list[tuple[str,tk.Entry]] = []

        _lbl(ef, "Indicator Distribution (fills automatically from bank)", bold=True).pack(
            anchor="w", pady=(10,2))
        tk.Label(ef,
                 text="Leave blank = outcome-only mode.  Fill any count = pair mode.",
                 bg=BG_L, fg=BG_M, font=FM).pack(anchor="w")
        self._ind_frame = tk.Frame(ef, bg=BG_L); self._ind_frame.pack(fill="x")

        # Params
        pf = tk.LabelFrame(p, text="Assembly Parameters",
                           bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        pf.grid(row=0, column=1, sticky="nsew", padx=(8,0), pady=4)
        self._nforms = self._le(pf, "Number of Forms:", 0, "4")
        tk.Label(pf, text="Difficulty Stage:", bg=BG_L, fg=TXD, font=FB).grid(
            row=1, column=0, sticky="w", padx=(0,8), pady=2)
        self._stage_var = tk.StringVar(value=list(MODE2_STAGES.keys())[0])
        self._stage_var.trace_add("write", self._on_stage)
        ttk.Combobox(pf, textvariable=self._stage_var,
                     values=list(MODE2_STAGES.keys()),
                     state="readonly", width=26).grid(row=1, column=1, sticky="w", pady=2)
        self._rng_lbl  = tk.Label(pf, text="", bg=BG_L, fg=TXD, font=FB)
        self._rng_ent  = _entry(pf, w=20); self._rng_ent.config(state="disabled")
        self._mean_lbl = tk.Label(pf, text="", bg=BG_L, fg=TXD, font=FB)
        self._mean_ent = _entry(pf, w=20); self._mean_ent.config(state="disabled")
        self._rng_lbl.grid(row=2, column=0, sticky="w")
        self._rng_ent.grid(row=2, column=1, sticky="w")
        self._mean_lbl.grid(row=3, column=0, sticky="w")
        self._mean_ent.grid(row=3, column=1, sticky="w")
        self._stats_var   = tk.BooleanVar(value=False)
        self._partial_var = tk.BooleanVar(value=True)
        self._reuse_var   = tk.BooleanVar(value=False)
        for row, (var, text) in enumerate([
            (self._stats_var,   "Enable discrimination filter (تمييز ≥ 0.50)"),
            (self._partial_var, "Allow partial fill (placeholder rows when short)"),
        ], start=4):
            tk.Checkbutton(pf, text=text, variable=var,
                           bg=BG_L, fg=TXD, font=FB).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=2)

        # Reuse row: checkbox + "Fill in:" label + spinbox on the same line
        reuse_frm = tk.Frame(pf, bg=BG_L)
        reuse_frm.grid(row=6, column=0, columnspan=2, sticky="w", pady=2)
        tk.Checkbutton(reuse_frm, text="Allow reuse across forms — never in same form",
                       variable=self._reuse_var,
                       bg=BG_L, fg=TXD, font=FB).pack(side="left")
        tk.Label(reuse_frm, text="   Fill in:",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._max_reuse_m2 = tk.Spinbox(
            reuse_frm, from_=2, to=10, width=4, font=FB,
            bg="#FFF", fg="#000", relief="solid", bd=1)
        self._max_reuse_m2.delete(0, "end"); self._max_reuse_m2.insert(0, "2")
        self._max_reuse_m2.pack(side="left", padx=4)
        tk.Label(reuse_frm, text="times max per question",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")

        self._sim_n = self._le(pf, "Simulation students:", 7, "3000")
        self._seed  = self._le(pf, "RNG Seed (blank=random):", 8, "")
        _btn(pf, "Save & Preview Settings", self._save_settings,
             bg=BG_D).grid(row=9, column=0, columnspan=2, pady=8)

    def _on_stage(self, *_):
        is_m = "Manually" in self._stage_var.get()
        if is_m:
            self._rng_lbl.config(text="Range (e.g. 0.2, 0.8):")
            self._rng_ent.config(state="normal")
            self._mean_lbl.config(text="Mean range (e.g. 0.5, 0.6):")
            self._mean_ent.config(state="normal")
        else:
            self._rng_lbl.config(text=""); self._rng_ent.config(state="disabled")
            self._mean_lbl.config(text=""); self._mean_ent.config(state="disabled")

    def _save_settings(self):
        try: n_forms = int(self._nforms.get())
        except: n_forms = "?"
        stage = self._stage_var.get()
        crit  = MODE2_STAGES.get(stage, {})
        # detect mode
        any_indicator = any(
            any(e.get().strip().isdigit() and int(e.get()) > 0
                for e in ind_dict.values())
            for _, ind_dict in self._indicator_entries
        )
        mode_str = "Pair mode (الناتج + المؤشر)" if any_indicator else "Outcome-only mode (الناتج)"
        lines = [
            f"Assembly mode   : {mode_str}",
            f"Stage           : {stage}",
            f"Diff Range      : {crit.get('Range', 'manual')}",
            f"Mean Target     : {crit.get('Mean', 'manual')}",
            f"Forms           : {n_forms}",
            f"Disc filter     : {'Enabled (≥0.5)' if self._stats_var.get() else 'Disabled'}",
            f"Partial fill    : {'Enabled' if self._partial_var.get() else 'Disabled'}",
            f"Reuse           : {'Enabled — max ' + self._max_reuse_m2.get() + 'x per question' if self._reuse_var.get() else 'Disabled'}",
        ]
        mb.showinfo("Settings Preview", "\n".join(lines))

    # ── Bins tab ──────────────────────────────────────────────────────────────
    def _build_bins_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "Manual Difficulty Bin Distribution", bold=True).pack(anchor="w", pady=(0,4))
        tk.Label(p,
                 text="Enter the number of questions per difficulty range for each outcome/indicator pair.",
                 bg=BG_L, fg=BG_M, font=FB).pack(anchor="w")
        self._use_bins = tk.BooleanVar(value=False)
        tk.Checkbutton(p, text="Enable manual bin distribution",
                       variable=self._use_bins, bg=BG_L, fg=TXD, font=FH).pack(
            anchor="w", pady=6)
        self._bins_inner = tk.Frame(p, bg=BG_L); self._bins_inner.pack(fill="x")

    def _populate_bins(self, df):
        for w in self._bins_inner.winfo_children(): w.destroy()
        self._bins_map.clear()
        pairs = df.groupby("_pair_key", sort=False).agg(
            الناتج=("الناتج","first"), المؤشر=("المؤشر","first")
        ).reset_index(drop=True)
        for _, row in pairs.iterrows():
            pk = str(row["الناتج"]) + "||" + str(row["المؤشر"])
            lf = tk.LabelFrame(
                self._bins_inner,
                text=f"Outcome: {row['الناتج']}  /  Indicator: {row['المؤشر']}",
                bg=BG_L, fg=BG_D, font=FM, padx=6, pady=4)
            lf.pack(fill="x", pady=2)
            entries = []
            for rw in range(2):
                frm = tk.Frame(lf, bg=BG_L); frm.pack(fill="x")
                for bi in range(5):
                    idx = rw*5+bi
                    tk.Label(frm, text=BIN_LABELS[idx]+":", bg=BG_L,
                             fg=TXD, font=FM, width=9).grid(row=0, column=bi*2, padx=2)
                    e = _entry(frm, w=4); e.grid(row=0, column=bi*2+1, padx=2)
                    entries.append(e)
            self._bins_map[pk] = entries

    # ── Generate tab ──────────────────────────────────────────────────────────
    def _build_generate_tab(self, p):
        p.configure(bg=BG_L)
        tk.Label(p, text="Generate Forms & Run 3PL Analysis",
                 bg=BG_L, fg=TXD, font=FH).pack(anchor="w", padx=20, pady=(14,4))
        bf = tk.Frame(p, bg=BG_L, pady=6); bf.pack(padx=20, anchor="w")
        self._gbtn = _btn(bf, "▶  Generate & Analyse", self._run, bg=self.COLOR)
        self._gbtn.config(font=("Segoe UI",11,"bold"), pady=8, padx=20)
        self._gbtn.pack(side="left", padx=(0,12))
        self._pv = tk.DoubleVar(value=0)
        ttk.Progressbar(bf, variable=self._pv, maximum=100,
                        length=320, mode="determinate").pack(side="left")
        lf = tk.LabelFrame(p, text="Log", bg=BG_L, fg=BG_D,
                           font=FH, padx=8, pady=8)
        lf.pack(fill="both", expand=True, padx=20, pady=8)
        self._logbox = _log_widget(lf)

    # ── Results tab ───────────────────────────────────────────────────────────
    def _build_results_tab(self, p):
        p.configure(bg=BG_L, padx=20, pady=16)
        _lbl(p, "Results Summary", bold=True).pack(anchor="w", pady=(0,8))
        self._rtree = _treeview(p,
            ["Form","Items","Mean Diff","Min","Max","Avg a","Avg c","Alpha","Item Corr."],
            [90, 60, 100, 60, 60, 75, 65, 100, 100])
        bf = tk.Frame(p, bg=BG_L); bf.pack(anchor="w", pady=8)
        _btn(bf, "Open Output Folder", self._open_folder, bg=self.COLOR).pack(side="left")
        _btn(bf, "Refresh", self._refresh_results, bg=BG_D).pack(side="left", padx=6)

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

    # ── Run logic ─────────────────────────────────────────────────────────────
    def _run(self):
        df = self._filtered_df or self.bank_df
        if df is None:
            mb.showerror("Error", "Please upload a question bank first."); return
        try: params, dcol = self._collect(df)
        except ValueError as e: mb.showerror("Configuration Error", str(e)); return
        self._gbtn.config(state="disabled"); self._pv.set(0)
        self._logbox.config(state="normal"); self._logbox.delete("1.0","end")
        self._logbox.config(state="disabled")
        threading.Thread(target=self._gen_t, args=(df, params, dcol), daemon=True).start()

    def _collect(self, df) -> tuple:
        def ii(e, n):
            try: return int(e.get().strip())
            except: raise ValueError(f"Invalid value for '{n}'")

        n_forms = ii(self._nforms, "Number of Forms")
        stage   = self._stage_var.get()
        crit    = MODE2_STAGES.get(stage, {})
        diff_range = crit.get("Range") or (0.0, 1.0)
        mean_range = crit.get("Mean")
        if "Manually" in stage:
            try:
                parts = [float(x) for x in self._rng_ent.get().split(",")]
                if len(parts) == 2: diff_range = tuple(parts)
            except: pass
            try:
                parts = [float(x) for x in self._mean_ent.get().split(",")]
                if len(parts) == 2: mean_range = tuple(parts)
            except: pass

        dcol = "difficulty"

        # Detect assembly mode:
        # If ANY indicator entry has a number → pair mode
        # Otherwise → outcome-only mode
        indicator_targets: dict[tuple[str,str], int] = {}
        for outcome, ind_dict in self._indicator_entries:
            for ind, ent in ind_dict.items():
                try: cnt = int(ent.get().strip())
                except: cnt = 0
                if cnt > 0:
                    indicator_targets[(outcome, ind)] = cnt

        pair_mode = len(indicator_targets) > 0

        slots: list[SlotDef] = []

        if pair_mode:
            for (outcome, ind), cnt in indicator_targets.items():
                pk = outcome + "||" + ind
                bins = []
                if self._use_bins.get() and pk in self._bins_map:
                    for bi, e in enumerate(self._bins_map[pk]):
                        try: bc = int(e.get())
                        except: bc = 0
                        if bc > 0:
                            bins.append(BinDef(bi*0.1, (bi+1)*0.1, bc))
                filt = {"الناتج": outcome}
                if ind not in ("", "nan", "None"):
                    filt["المؤشر"] = ind
                slots.append(SlotDef(label=f"{outcome} / {ind}",
                                      filters=filt, count=cnt, bins=bins))
        else:
            # outcome-only mode: use outcome rows (from template or bank)
            for outcome in df["الناتج"].dropna().unique():
                # default count: 5, user can override via template
                cnt = 5
                for (o, ind_dict) in self._indicator_entries:
                    if o == outcome:
                        pass   # no indicator counts
                pk = outcome + "||__all__"
                bins = []
                if self._use_bins.get() and pk in self._bins_map:
                    for bi, e in enumerate(self._bins_map[pk]):
                        try: bc = int(e.get())
                        except: bc = 0
                        if bc > 0:
                            bins.append(BinDef(bi*0.1, (bi+1)*0.1, bc))
                slots.append(SlotDef(label=str(outcome),
                                      filters={"الناتج": outcome},
                                      count=cnt, bins=bins))

        if not slots:
            raise ValueError(
                "No assembly slots defined.\n"
                "Upload the bank and fill in at least one indicator count "
                "(pair mode) or ensure outcomes are detected.")

        seed_s = self._seed.get().strip()
        try:
            max_reuse_val = max(2, int(self._max_reuse_m2.get()))
        except ValueError:
            max_reuse_val = 2
        params = AssemblyParams(
            slots=slots,
            diff_range=diff_range,
            diff_mean_range=mean_range,
            min_discrimination=0.5 if self._stats_var.get() else None,
            allow_partial_fill=self._partial_var.get(),
            allow_reuse=self._reuse_var.get(),
            max_reuse=max_reuse_val,
            max_retries=100,
            rng_seed=int(seed_s) if seed_s else None,
        )
        self._n_forms_last = n_forms
        return params, dcol

    def _gen_t(self, df, params, dcol):
        try:
            n = self._n_forms_last
            names = [f"Form_{i+1}" for i in range(n)]
            self._lg(f"Starting assembly: {n} forms…")
            results = assemble_forms(df, params, dcol, n, names)
            self._q.put(("prog", 40))

            for (form, mean_d, warns), name in zip(results, names):
                for w in warns: self._lg(f"  ⚠ {w}", "yellow")
                self._lg(f"  {name}: {len(form)} items, mean difficulty = {mean_d:.3f}")

            self._lg("Running 3PL psychometric analysis…")
            analyses = [
                analyse_form(r[0], nm, dcol, "تمييز", "التخمين",
                             3000, None)
                for r, nm in zip(results, names)
            ]
            self._q.put(("prog", 75))
            for fa in analyses:
                if fa:
                    self._lg(f"  {fa.name}: Alpha = {fa.alpha:.3f}, "
                             f"Item Corr. = {fa.item_corr:.3f}")

            self._lg("Writing output files…")
            base   = self.source_path.parent
            prefix = self.source_path.stem
            fp = base / f"{prefix}_Forms.xlsx"
            ap = base / f"{prefix}_3PL_Analysis.xlsx"
            rp = base / f"{prefix}_Remaining_Questions.xlsx"

            usage: dict[str, int] = {}
            for form, _, _ in results:
                if "QuestionID" in form.columns:
                    for qid in form["QuestionID"].dropna().astype(str):
                        usage[qid] = usage.get(qid, 0) + 1

            write_forms_workbook(
                results, names, analyses,
                dcol, "المجال", "تمييز", "التخمين",
                usage, fp, mode=2)
            valid = [fa for fa in analyses if fa]
            if valid: write_analysis_workbook(valid, ap)

            used_ids = set()
            for form, _, _ in results:
                if "QuestionID" in form.columns:
                    used_ids.update(form["QuestionID"].dropna().astype(str).tolist())
            remaining = df[~df["QuestionID"].astype(str).isin(used_ids)].copy()
            remaining.drop(columns=[c for c in remaining.columns if c.startswith("_")],
                           errors="ignore").to_excel(str(rp), index=False, engine="openpyxl")

            self._q.put(("prog", 100))
            self._lg(f"\n  📄 Forms           : {fp}")
            self._lg(f"  📊 3PL Analysis    : {ap}")
            self._lg(f"  📋 Remaining Qs    : {rp}")
            self._lg("\n✓ Generation complete!", "lime")
            self._q.put(("done", valid))

        except Exception as e:
            self._lg(f"\n✗ Error: {e}", "red")
            log.exception("Mode2 generation failed")
            self._q.put(("fail",))

# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    LauncherWindow().mainloop()
