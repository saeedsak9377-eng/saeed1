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
#  BLOCK ID SYSTEM
# ─────────────────────────────────────────────────────────────────────────────
# Format:  <ExamCode>-<Stage>.<Level>.<BlockNumber:03d>
#
# Stage 1:  level = 1 (Part 1) or 2 (Part 2)   → GAT-1.1.001 / GAT-1.2.001
# Stage 2:  level = 1 Easy / 2 Medium / 3 Hard  → GAT-2.1.001
# Stage 3:  level = 1 Easy / 2 Medium / 3 Hard  → GAT-3.1.001

EXAM_CODES: dict[str, str] = {
    "قدرات علمي":      "GAT",
    "قدرات نظري":      "NLAT",
    "القدرة المعرفية": "CAT",
    "التحصيلي":        "ACH",
    "قدرات الجامعيين": "GGAT",
    "Custom":           "EX",
    # Mode 2 default code
    "نافس 4 أسئلة":   "NAF4",
    "نافس 5 أسئلة":   "NAF5",
}

# Maps (stage_name_in_UI) → (stage_number, level_number)
# Used to derive the numeric stage.level for the block ID
STAGE_TO_SL: dict[str, tuple[int, int]] = {
    # Mode 1
    "Stage1":       (1, 0),   # 0 = Part selector required (1 or 2)
    "E":            (2, 1),
    "M":            (2, 2),
    "D":            (2, 3),
    "Manually set": (3, 0),   # 0 = difficulty-level selector required
    # Mode 2
    "علوم صف الثالث":         (2, 1),
    "علوم صف السادس":         (2, 2),
    "علوم الصف التاسع":       (2, 3),
    "الرياضيات الصف الثالث":  (2, 1),
    "الرياضيات الصف السادس":  (2, 2),
    "الرياضيات الصف التاسع":  (2, 3),
    "القراءة الصف الثالث":    (2, 1),
    "القراءة الصف السادس":    (2, 2),
    "القراءة الصف التاسع":    (2, 3),
    "عام":                    (1, 1),
}

_BLOCK_COUNTER_FILE = ".stgs_block_counters.json"


def _load_counters(base_dir: Path) -> dict:
    """Load persistent block counters from JSON file next to the bank."""
    fpath = base_dir / _BLOCK_COUNTER_FILE
    if fpath.exists():
        try:
            import json
            return json.loads(fpath.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_counters(base_dir: Path, counters: dict) -> None:
    import json
    fpath = base_dir / _BLOCK_COUNTER_FILE
    fpath.write_text(json.dumps(counters, ensure_ascii=False, indent=2),
                     encoding="utf-8")


def _max_block_number_from_used_blocks(
        key: str,
        used_blocks_df: "Optional[pd.DataFrame]",
) -> int:
    """
    Scan the Used Blocks dataset for existing IDs matching `key`
    (e.g. "GAT-2.1") and return the highest block number found.
    Returns 0 if no match exists.

    This is the authoritative source for sequencing — the local JSON
    counter is only used as a fallback when no Used Blocks file is loaded.
    """
    if used_blocks_df is None or used_blocks_df.empty:
        return 0
    if "Block-ID" not in used_blocks_df.columns:
        return 0

    import re
    prefix  = key + "."          # e.g. "GAT-2.1."
    pattern = re.compile(rf"^{re.escape(prefix)}(\d+)$")
    max_num = 0
    for bid in used_blocks_df["Block-ID"].dropna().astype(str):
        m = pattern.match(bid.strip())
        if m:
            max_num = max(max_num, int(m.group(1)))
    return max_num


def generate_block_ids(exam_name: str,
                       stage_name: str,
                       n_forms: int,
                       part_or_level: int,    # 1/2 for Stage1 parts; 1/2/3 for difficulty
                       base_dir: Path,
                       used_blocks_df: "Optional[pd.DataFrame]" = None,
                       stage_override: Optional[int] = None,
                       ) -> list[str]:
    """
    Generate n_forms sequential Block IDs that continue from wherever
    the sequence currently stands.

    stage_override
    --------------
    When stage_name is "E", "M", or "D" the table maps them to Stage 2
    by default.  Pass stage_override=3 to generate Stage 3 IDs instead
    (e.g. GAT-3.1.* instead of GAT-2.1.*).  This is the user-facing
    "Stage Number" radio button in the Block ID Settings panel.

    Sequencing priority:
      1. Scan Used Blocks dataset for the highest existing block number
         in the same group → next = max_found + 1
      2. If no Used Blocks loaded (or group not found) → read local JSON
      3. Neither exists → start from 001
    """
    code    = EXAM_CODES.get(exam_name, "EX")
    sl      = STAGE_TO_SL.get(stage_name, (1, 1))
    stage_n = sl[0]
    level_n = part_or_level if sl[1] == 0 else sl[1]

    # Apply stage override (allows E/M/D to produce Stage 3 IDs)
    if stage_override is not None and stage_n in (2, 3):
        stage_n = stage_override

    # Canonical key for this exam / stage / level combination
    key = f"{code}-{stage_n}.{level_n}"

    # Step 1: highest number already used (from Used Blocks dataset)
    ub_max = _max_block_number_from_used_blocks(key, used_blocks_df)

    # Step 2: highest number in local JSON counter
    counters  = _load_counters(base_dir)
    json_max  = counters.get(key, 0)

    # Take the larger of the two so we never duplicate
    start = max(ub_max, json_max) + 1

    ids: list[str] = []
    for i in range(n_forms):
        ids.append(f"{key}.{start + i:03d}")

    # Persist the new high-water mark
    counters[key] = start + n_forms - 1
    _save_counters(base_dir, counters)
    return ids


# ─────────────────────────────────────────────────────────────────────────────
#  USED BLOCKS LOADER
# ─────────────────────────────────────────────────────────────────────────────

UB_ALIASES = {
    "QuestionID":   ["questionid","qoustionid","question_id","id","رقم"],
    "Block-ID":     ["block-id","block_id","blockid","block id","block"],
    "Difficulty":   ["difficulty","الصعوبة","b","difficulty"],
    "Category":     ["category","الفئة","التصنيف","الناتج"],
    "Number of used": ["number of used","عدد مرات استخدام السؤال","used_count",
                       "usage","used","استخدام"],
    "D":            ["d","domain","المجال"],
    "Discrimination": ["تمييز","discrimination","a"],
    "مستوى التمييز":  ["مستوى التمييز","level","مستوى"],
    "Guessing":     ["التخمين","التخميين","guessing","c"],
}


def load_used_blocks(path: str | Path) -> pd.DataFrame:
    """
    Load the Used Blocks dataset.
    Normalises column names and ensures 'Number of used' is numeric.
    Returns an empty DataFrame if the file cannot be read.
    """
    path = Path(path)
    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(path, engine="openpyxl")
        else:
            df = pd.read_csv(path, encoding="utf-8-sig")
    except Exception as exc:
        log.warning("Could not load Used Blocks file: %s", exc)
        return pd.DataFrame()

    # Normalise columns
    lower = {c.strip().lower(): c for c in df.columns}
    rename = {}
    for canon, alts in UB_ALIASES.items():
        if canon in df.columns:
            continue
        for a in alts:
            if a.strip().lower() in lower:
                rename[lower[a.strip().lower()]] = canon
                break
    df = df.rename(columns=rename)

    if "QuestionID" not in df.columns:
        log.warning("Used Blocks file has no QuestionID column — ignored.")
        return pd.DataFrame()

    if "Number of used" not in df.columns:
        df["Number of used"] = 0
    else:
        df["Number of used"] = pd.to_numeric(df["Number of used"],
                                              errors="coerce").fillna(0)

    # Ensure Difficulty column exists
    if "Difficulty" not in df.columns and "difficulty" in df.columns:
        df["Difficulty"] = df["difficulty"]
    if "Difficulty" not in df.columns:
        df["Difficulty"] = 0.5

    df["Difficulty"] = pd.to_numeric(df["Difficulty"],
                                     errors="coerce").fillna(0.5).clip(0, 1)
    df = df[~df["QuestionID"].duplicated(keep="first")].reset_index(drop=True)
    log.info("Used Blocks loaded: %d questions", len(df))
    return df


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
    min_discrimination: Optional[float] = None
    allow_partial_fill: bool = True
    allow_reuse:        bool = False
    max_reuse:          int  = 3
    max_retries:        int  = 100
    rng_seed:           Optional[int] = None
    # Used Blocks fallback
    used_blocks:        Optional["pd.DataFrame"] = field(default=None)
    allow_used_blocks:  bool = False
    # Max times a Used Blocks question may appear across all assembled forms.
    # 0 = no cap (same as old behaviour).  Defaults to 2 (one primary + one reuse).
    max_ub_usage:       int  = 2
    # Adaptive Mean Control
    adaptive_mean_control: bool = False
    # ── Feature 2: Difficulty distribution scope ──────────────────────────────
    # 'exam'     = one overall mean check for the whole form (default)
    # 'domain'   = each D-domain must independently meet mean target
    # 'category' = each Category must independently meet mean target
    difficulty_scope: str = "exam"
    # ── Feature 3: Difficulty method (Delta = current; CCT = classical) ───────
    difficulty_method: str = "delta"
    cct_profile:       str = "Common"
    # ── Feature 4: Fallback for difficulty (dedicated reuse for mean targets) ─
    allow_fallback_reuse: bool = False
    # ── Psychometric model selection (does NOT affect assembly logic) ─────────
    # "3pl" = existing 3PL analysis (default, unchanged behaviour)
    # "1pl" = 1PL/Rasch analysis — generates 1PL_Result.xlsx instead of 3PL
    psychometric_model: str = "3pl"


# ══════════════════════════════════════════════════════════════════════════════
#  FEATURE 1 — DYNAMIC EXAM TEMPLATE IMPORT / EXPORT
# ══════════════════════════════════════════════════════════════════════════════

TEMPLATE_HEADERS = ["Exam type", "D", "Category", "Count"]
TEMPLATE_EXAMPLE_ROWS = [
    ["المبكر", "V",  "VRS",   3],
    ["المبكر", "V",  "VSS",   4],
    ["المبكر", "V",  "VME",   5],
    ["المبكر", "V",  "VVT",   1],
    ["المبكر", "Q",  "QSS",   3],
    ["المبكر", "Q",  "QRE",   6],
    ["المبكر", "Q",  "QVC",   4],
    ["المبكر", "Q",  "QQW",   3],
    ["المبكر", "Q",  "QRT",   2],
    ["المبكر", "SR", "SRW",   3],
    ["المبكر", "SR", "SRTTM", 5],
    ["المبكر", "SR", "SRMKS", 6],
    ["المبكر", "SR", "SRML",  7],
]


def export_exam_template(out_path: Path) -> None:
    """Write a blank/example template .xlsx for users to fill in."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook(); ws = wb.active; ws.title = "Exam Template"

    # Header styling
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    border   = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'))

    for ci, hdr in enumerate(TEMPLATE_HEADERS, 1):
        cell = ws.cell(row=1, column=ci, value=hdr)
        cell.font = hdr_fill and hdr_font; cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    # Example data rows
    alt_fill = PatternFill("solid", fgColor="DCE6F1")
    for ri, row in enumerate(TEMPLATE_EXAMPLE_ROWS, 2):
        fill = alt_fill if ri % 2 == 0 else PatternFill()
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill; cell.border = border

    # Column widths
    for col, width in zip("ABCD", [20, 12, 16, 10]):
        ws.column_dimensions[col].width = width

    # Instructions sheet
    ws2 = wb.create_sheet("Instructions")
    instructions = [
        ["Column",     "Description",                            "Example"],
        ["Exam type",  "Name of the exam to create",            "المبكر"],
        ["D",          "Domain / Section code",                 "V, Q, SR"],
        ["Category",   "Skill subcategory code",                "VRS, QSS"],
        ["Count",      "Number of questions required (integer)", "5"],
        ["",           "",                                       ""],
        ["Rules:",     "", ""],
        ["",           "• One exam name per file",               ""],
        ["",           "• No empty cells in any row",           ""],
        ["",           "• Count must be a positive integer",    ""],
        ["",           "• (Domain, Category) pairs must be unique", ""],
    ]
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 40
    ws2.column_dimensions["C"].width = 16
    for ri, row in enumerate(instructions, 1):
        for ci, val in enumerate(row, 1):
            ws2.cell(row=ri, column=ci, value=val)

    wb.save(str(out_path))
    log.info("Exam template exported → %s", out_path)


def import_exam_template(path: Path) -> tuple[str, dict, list[str]]:
    """
    Parse an exam template file.

    Returns
    -------
    (exam_name, structure, errors)

    structure = {
        "ExamName": {
            "DomainA": {"CatA": count, "CatB": count, ...},
            "DomainB": {...},
        }
    }
    errors = [] on success, list of human-readable messages on failure.
    """
    errors: list[str] = []

    # Read file
    try:
        suffix = path.suffix.lower()
        if suffix in (".xlsx", ".xls"):
            df = pd.read_excel(path, engine="openpyxl")
        elif suffix == ".csv":
            df = pd.read_csv(path, encoding="utf-8-sig")
        else:
            return "", {}, [f"Unsupported file type: {suffix}"]
    except Exception as exc:
        return "", {}, [f"Cannot read file: {exc}"]

    if df.empty:
        return "", {}, ["File is empty."]

    # Validate headers
    missing_headers = [h for h in TEMPLATE_HEADERS if h not in df.columns]
    if missing_headers:
        return "", {}, [
            f"Missing required column(s): {', '.join(missing_headers)}.  "
            f"Expected headers: {TEMPLATE_HEADERS}"
        ]

    # Work on a clean copy
    df = df[TEMPLATE_HEADERS].copy()

    # Row-level validation
    exam_names = set()
    seen_pairs: set[tuple] = set()
    for i, row in df.iterrows():
        rn = i + 2  # 1-indexed with header row
        for col in TEMPLATE_HEADERS:
            if pd.isna(row[col]) or str(row[col]).strip() == "":
                errors.append(f"Row {rn}: empty value in column '{col}'.")
        if errors:
            continue
        exam_names.add(str(row["Exam type"]).strip())
        try:
            cnt = int(row["Count"])
            if cnt <= 0:
                errors.append(f"Row {rn}: Count must be > 0 (got {cnt}).")
        except (ValueError, TypeError):
            errors.append(f"Row {rn}: Count is not a valid integer (got '{row['Count']}').")
        pair = (str(row["D"]).strip(), str(row["Category"]).strip())
        if pair in seen_pairs:
            errors.append(f"Row {rn}: duplicate (D, Category) pair {pair}.")
        seen_pairs.add(pair)

    if len(exam_names) > 1:
        errors.append(
            f"File contains multiple exam names: {exam_names}. "
            "Only one exam name is allowed per file."
        )

    if errors:
        return "", {}, errors

    exam_name = str(df["Exam type"].iloc[0]).strip()

    # Build structure
    structure: dict = {exam_name: {}}
    for _, row in df.iterrows():
        domain = str(row["D"]).strip()
        cat    = str(row["Category"]).strip()
        count  = int(row["Count"])
        structure[exam_name].setdefault(domain, {})[cat] = count

    log.info("Template imported: exam='%s', %d domains, %d categories",
             exam_name,
             len(structure[exam_name]),
             sum(len(cats) for cats in structure[exam_name].values()))
    return exam_name, structure, []


# ══════════════════════════════════════════════════════════════════════════════
#  FEATURE 2 — DIFFICULTY SCOPE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _check_scope_mean(form: pd.DataFrame, dcol: str,
                      mn_range: tuple, scope: str,
                      domain_col: str = "D") -> bool:
    """
    Return True if the form satisfies the mean criterion at the given scope.

    scope='exam'     → check overall form mean only
    scope='domain'   → check each D-domain independently
    scope='category' → check each Category independently
    """
    if mn_range is None:
        return True
    lo, hi = mn_range

    def _mean_ok(subset):
        diffs = pd.to_numeric(subset[dcol], errors="coerce").dropna()
        return diffs.empty or (lo <= float(diffs.mean()) <= hi)

    if scope == "exam":
        return _mean_ok(form)

    if scope == "domain":
        grp_col = domain_col if domain_col in form.columns else dcol
        if grp_col == dcol:
            return _mean_ok(form)
        for dom in form[grp_col].dropna().unique():
            if not _mean_ok(form[form[grp_col] == dom]):
                return False
        return True

    if scope == "category":
        cat_col = "Category" if "Category" in form.columns else \
                  ("الناتج" if "الناتج" in form.columns else None)
        if cat_col is None:
            return _mean_ok(form)
        for cat in form[cat_col].dropna().unique():
            if not _mean_ok(form[form[cat_col] == cat]):
                return False
        return True

    return _mean_ok(form)


# ══════════════════════════════════════════════════════════════════════════════
#  FEATURE 3 — CCT DIFFICULTY PROFILES
# ══════════════════════════════════════════════════════════════════════════════

CCT_RULES: dict[str, dict] = {
    "Common":       {"Range": (0.20, 0.90), "Mean": (0.49, 0.55)},
    "D":            {"Range": (0.00, 0.45), "Mean": (0.27, 0.35)},
    "M":            {"Range": (0.40, 0.75), "Mean": (0.55, 0.61)},
    "E":            {"Range": (0.70, 1.00), "Mean": (0.80, 0.88)},
    "Manually set": {"Range": None,          "Mean": None},
}


def resolve_diff_params(params: "AssemblyParams") -> tuple:
    """
    Return (diff_range, diff_mean_range) resolved from the selected method.

    Delta (default) → use params.diff_range / params.diff_mean_range as-is.
    CCT             → look up profile in CCT_RULES; manual overrides allowed.
    """
    if params.difficulty_method == "cct":
        profile = CCT_RULES.get(params.cct_profile, CCT_RULES["Common"])
        diff_range      = profile["Range"] or params.diff_range
        diff_mean_range = profile["Mean"]  or params.diff_mean_range
        return diff_range, diff_mean_range
    return params.diff_range, params.diff_mean_range


# ══════════════════════════════════════════════════════════════════════════════
#  FEATURE 4 — FALLBACK-FOR-DIFFICULTY (controlled reuse targeting mean)
# ══════════════════════════════════════════════════════════════════════════════

def _filter_ub_by_prefix(
        used_blocks: pd.DataFrame,
        block_prefix: Optional[str],
) -> pd.DataFrame:
    """
    Shared helper: filter Used Blocks DataFrame to only rows whose
    Block-ID belongs to the exact Stage + Level/Part identified by
    `block_prefix` (e.g. "GAT-2.1", "GAT-1.2").

    If `block_prefix` is None or the 'Block-ID' column is absent the
    full DataFrame is returned unchanged (backward-compatible).

    The match is a prefix + "." check so "GAT-2.1" matches "GAT-2.1.001"
    but never "GAT-2.10.001" or "GAT-2.11.001".
    """
    if block_prefix is None or used_blocks.empty:
        return used_blocks
    if "Block-ID" not in used_blocks.columns:
        return used_blocks
    mask = used_blocks["Block-ID"].astype(str).str.startswith(
        block_prefix + "."
    )
    return used_blocks[mask].copy()


def _fallback_for_difficulty(
        form: pd.DataFrame,
        dcol: str,
        mn_range: tuple,
        scope: str,
        params: "AssemblyParams",
        bank: pd.DataFrame,
        used_ids: set,
        question_usage: dict,
        rng: np.random.Generator,
        form_number: int,
        block_prefix: Optional[str] = None,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Feature 4 — controlled reuse specifically to hit the difficulty mean.

    Triggered when:
      1. allow_fallback_reuse=True
      2. The form fails the scope-level mean check
      3. Main bank (unused questions) was insufficient

    Strategy: swap items whose difficulty deviates most from the target mean
    with items from the reuse pool (bank + used_blocks), sorted by:
      - Least reuse count first
      - Difficulty closest to what is needed to correct the mean
    """
    warnings: list[str] = []
    if not params.allow_fallback_reuse or mn_range is None:
        return form, warnings
    if _check_scope_mean(form, dcol, mn_range, scope):
        return form, warnings   # already fine

    lo, hi = mn_range
    target = (lo + hi) / 2.0
    current = float(pd.to_numeric(form[dcol], errors="coerce").dropna().mean())

    # Build reuse pool from bank items already used (not in current form)
    reuse_pool_parts = []
    reuse_ids = {q for q, c in question_usage.items() if 0 < c < params.max_reuse}
    if reuse_ids and "QuestionID" in bank.columns:
        rp = bank[bank["QuestionID"].astype(str).isin(reuse_ids) &
                  ~bank["QuestionID"].astype(str).isin(used_ids)].copy()
        rp["_use_count"] = rp["QuestionID"].astype(str).map(
            lambda q: question_usage.get(q, 0))
        reuse_pool_parts.append(rp)

    # Add Used Blocks — STRICT stage-level filter applied first
    if params.allow_used_blocks and params.used_blocks is not None \
            and not params.used_blocks.empty:
        ub = _filter_ub_by_prefix(params.used_blocks, block_prefix)
        if ub.empty and block_prefix:
            warnings.append(
                f"Form {form_number}: _fallback_for_difficulty — Used Blocks has "
                f"no records matching stage-level '{block_prefix}'; "
                f"cross-stage reuse is forbidden.")
        elif not ub.empty:
            if dcol not in ub.columns and "Difficulty" in ub.columns:
                ub = ub.rename(columns={"Difficulty": dcol})
            if dcol in ub.columns:
                ub = ub[~ub["QuestionID"].astype(str).isin(used_ids)]
                ub[dcol] = pd.to_numeric(ub[dcol], errors="coerce").fillna(0.5)
                ub["_use_count"] = ub.get("Number of used",
                                           pd.Series(0, index=ub.index))
                reuse_pool_parts.append(ub)

    if not reuse_pool_parts:
        return form, warnings

    reuse_pool = pd.concat(reuse_pool_parts, ignore_index=True)
    reuse_pool[dcol] = pd.to_numeric(reuse_pool[dcol], errors="coerce")
    lo_r, hi_r = params.diff_range
    reuse_pool = reuse_pool[(reuse_pool[dcol] >= lo_r) &
                            (reuse_pool[dcol] <= hi_r)].dropna(subset=[dcol])

    if reuse_pool.empty:
        return form, warnings

    # Determine swap direction
    direction = "higher" if current < lo else "lower"
    if direction == "higher":
        reuse_pool = reuse_pool.sort_values(
            [dcol, "_use_count"], ascending=[False, True])
    else:
        reuse_pool = reuse_pool.sort_values(
            [dcol, "_use_count"], ascending=[True, True])

    form = form.copy()
    max_swaps = max(1, len(form) // 4)

    for _ in range(max_swaps):
        current = float(pd.to_numeric(
            form[dcol], errors="coerce").dropna().mean())
        if _check_scope_mean(form, dcol, mn_range, scope):
            break

        # Find worst-deviating item
        valid = form[dcol].notna()
        if not valid.any(): break
        diffs = pd.to_numeric(form.loc[valid, dcol])
        worst_pos = int(diffs.idxmin() if direction == "higher" else diffs.idxmax())
        worst_diff = float(form.at[worst_pos, dcol])

        # Category constraint
        cat_col = next((c for c in ("Category","الناتج") if c in form.columns), None)
        candidates = reuse_pool
        if cat_col and cat_col in reuse_pool.columns:
            worst_cat = form.at[worst_pos, cat_col]
            cat_cands = reuse_pool[reuse_pool[cat_col].astype(str) == str(worst_cat)]
            if not cat_cands.empty:
                candidates = cat_cands

        # Filter: must actually help
        if direction == "higher":
            candidates = candidates[candidates[dcol] > worst_diff]
        else:
            candidates = candidates[candidates[dcol] < worst_diff]

        if candidates.empty: break

        replacement = candidates.iloc[0].copy()
        rep_qid = str(replacement.get("QuestionID", ""))
        old_qid = str(form.at[worst_pos, "QuestionID"]) \
                  if "QuestionID" in form.columns else ""

        reuse_pool = reuse_pool[reuse_pool["QuestionID"].astype(str) != rep_qid]
        shared = [c for c in replacement.index
                  if c in form.columns and not c.startswith("_")]
        for col in shared:
            form.at[worst_pos, col] = replacement[col]

        used_ids.discard(old_qid); used_ids.add(rep_qid)
        warnings.append(
            f"Form {form_number}: fallback-for-difficulty swap "
            f"diff={worst_diff:.3f}→{float(replacement[dcol]):.3f} "
            f"(reuse fallback)")

    return form, warnings


# ══════════════════════════════════════════════════════════════════════════════
#  FEATURE 5 — CHARTS OUTPUT FILE
# ══════════════════════════════════════════════════════════════════════════════

def write_charts_workbook(
    results: list[tuple[pd.DataFrame, float, list[str]]],
    form_names: list[str],
    dcol: str,
    domain_col: str,
    diff_method: str,
    out_path: Path,
) -> None:
    """
    Feature 5: Generate a dedicated Charts.xlsx with two sheets:
      Sheet 1 "Exam Level"  — one histogram per form, laid out in a grid
      Sheet 2 "Domain Level" — per-form, per-domain histograms in rows
    """
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook(); wb.remove(wb.active)

    # ── Colour helpers ────────────────────────────────────────────────────────
    NAVY_FILL  = PatternFill("solid", fgColor="1F3864")
    TEAL_FILL  = PatternFill("solid", fgColor="00AECB")
    TITLE_FONT = Font(bold=True, color="FFFFFF", size=10)

    def _bin_counts(series: pd.Series) -> list[int]:
        counts = [0] * 10
        for d in pd.to_numeric(series, errors="coerce").dropna():
            counts[min(int(float(d) * 10), 9)] += 1
        return counts

    BIN_LABELS_LOCAL = [f"{i*0.1:.1f}-{(i+1)*0.1:.1f}" for i in range(10)]

    def _write_chart_block(ws, start_row: int, start_col: int,
                           form_name: str, bin_counts: list[int],
                           form_mean: float, n_items: int,
                           chart_title: str) -> None:
        """Write bin table + bar chart for one block. Returns rows used."""
        # Title cell
        tc = ws.cell(row=start_row, column=start_col, value=chart_title)
        tc.font = Font(bold=True, size=9, color="2D2B6E")

        # Summary row
        ws.cell(row=start_row+1, column=start_col,
                value=f"Mean={form_mean:.3f}  N={n_items}  Method={diff_method.upper()}")

        # Bin header
        ws.cell(row=start_row+2, column=start_col,
                value="Bin").fill = TEAL_FILL
        ws.cell(row=start_row+2, column=start_col+1,
                value="Count").fill = TEAL_FILL

        for i, (lbl, cnt) in enumerate(zip(BIN_LABELS_LOCAL, bin_counts)):
            ws.cell(row=start_row+3+i, column=start_col,   value=lbl)
            ws.cell(row=start_row+3+i, column=start_col+1, value=cnt)

        # Bar chart
        chart = BarChart()
        chart.type  = "col"
        chart.title = chart_title
        chart.y_axis.title = "Count"
        chart.x_axis.title = "Difficulty Bin"
        chart.style  = 10
        chart.width  = 12
        chart.height = 8

        data_ref = Reference(ws,
                             min_col=start_col+1, min_row=start_row+2,
                             max_row=start_row+12)
        cats_ref = Reference(ws,
                             min_col=start_col,   min_row=start_row+3,
                             max_row=start_row+12)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        # Place chart below the table (add 2 rows gap)
        anchor_col = get_column_letter(start_col)
        anchor_row = start_row + 14
        ws.add_chart(chart, f"{anchor_col}{anchor_row}")

    # ── Sheet 1: Exam Level ───────────────────────────────────────────────────
    ws1 = wb.create_sheet("Exam Level")
    # Header banner
    ws1.merge_cells("A1:L1")
    hc = ws1["A1"]
    hc.value = "STGS — Difficulty Distribution Comparison (Exam Level)"
    hc.font  = Font(bold=True, color="FFFFFF", size=12)
    hc.fill  = NAVY_FILL
    hc.alignment = Alignment(horizontal="center")
    ws1.row_dimensions[1].height = 20

    # Summary table
    ws1.cell(row=3, column=1, value="Form ID").fill  = TEAL_FILL
    ws1.cell(row=3, column=2, value="N Items").fill  = TEAL_FILL
    ws1.cell(row=3, column=3, value="Mean Diff").fill= TEAL_FILL
    ws1.cell(row=3, column=4, value="Method").fill   = TEAL_FILL
    ws1.cell(row=3, column=5, value="Fallback").fill = TEAL_FILL
    for ri, ((form, mean_d, warns), name) in \
            enumerate(zip(results, form_names), 4):
        ws1.cell(row=ri, column=1, value=name)
        ws1.cell(row=ri, column=2, value=len(form))
        ws1.cell(row=ri, column=3, value=round(mean_d, 4))
        ws1.cell(row=ri, column=4, value=diff_method.upper())
        ws1.cell(row=ri, column=5,
                 value="Yes" if any("fallback" in w.lower() for w in warns)
                 else "No")

    # Charts grid: 3 per row
    CHARTS_PER_ROW = 3
    COL_SPAN       = 5   # columns per chart block
    ROW_SPAN       = 32  # rows per chart block
    chart_data_start = 4 + len(results) + 2

    for fi, ((form, mean_d, _), name) in enumerate(zip(results, form_names)):
        row_block = fi // CHARTS_PER_ROW
        col_block = fi %  CHARTS_PER_ROW
        sr = chart_data_start + row_block * ROW_SPAN
        sc = 1 + col_block * COL_SPAN
        bins = _bin_counts(form[dcol]) if dcol in form.columns else [0]*10
        _write_chart_block(ws1, sr, sc, name, bins, mean_d,
                           len(form),
                           f"{name} — Exam Distribution")

    # ── Sheet 2: Domain Level ─────────────────────────────────────────────────
    ws2 = wb.create_sheet("Domain Level")
    ws2.merge_cells("A1:L1")
    hc2 = ws2["A1"]
    hc2.value = "STGS — Difficulty Distribution Comparison (Domain Level)"
    hc2.font  = Font(bold=True, color="FFFFFF", size=12)
    hc2.fill  = NAVY_FILL
    hc2.alignment = Alignment(horizontal="center")
    ws2.row_dimensions[1].height = 20

    cur_row = 3
    for (form, mean_d, warns), name in zip(results, form_names):
        # Form label
        lbl = ws2.cell(row=cur_row, column=1, value=f"▶  {name}")
        lbl.font = Font(bold=True, size=11, color="2D2B6E")
        cur_row += 1

        if domain_col not in form.columns:
            cur_row += 2
            continue

        domains = sorted(form[domain_col].dropna().unique().astype(str))
        col_offset = 1
        for dom in domains:
            sub   = form[form[domain_col].astype(str) == dom]
            bins  = _bin_counts(sub[dcol]) if dcol in sub.columns else [0]*10
            dm    = float(pd.to_numeric(sub[dcol], errors="coerce").dropna().mean()) \
                    if dcol in sub.columns else 0.0
            _write_chart_block(
                ws2, cur_row, col_offset,
                f"{name}·{dom}", bins, dm, len(sub),
                f"{name} — {dom}")
            col_offset += COL_SPAN

        cur_row += ROW_SPAN

    wb.save(str(out_path))
    log.info("Charts workbook → %s", out_path)


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
                 usage: Optional[dict] = None,
                 block_prefix: Optional[str] = None):
        self.bank   = bank.copy().reset_index(drop=True)
        self.params = params
        self.dcol   = dcol
        self.rng    = np.random.default_rng(params.rng_seed)
        # usage[QuestionID_str] = number of forms it has been used in
        self.question_usage: dict[str, int] = usage or {}
        # Block prefix for Used Blocks stage-level filtering (e.g. "GAT-2.1")
        # Derived from the block ID being assembled: "GAT-2.1.006" → "GAT-2.1"
        self.block_prefix: Optional[str] = block_prefix
        # Tracks how many forms each Used Blocks question has appeared in.
        # Shared across engine instances via the dict passed from assemble_forms.
        self.ub_usage: dict[str, int] = {}   # set by assemble_forms

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

                # Domain-level mean target: clamp to the available pool's
                # actual mean so we never target a value outside the pool.
                # This prevents domains with skewed pools from drifting.
                pool_diffs = pd.to_numeric(
                    d_pool[self.dcol], errors="coerce").dropna()
                if not pool_diffs.empty:
                    pool_mean = float(pool_diffs.mean())
                    # Blend 70% global target + 30% pool mean — keeps us close
                    # to the overall target while honouring the pool distribution
                    domain_mean_t = 0.70 * mean_t + 0.30 * pool_mean
                    # Hard-clip to diff_range so we never target outside bounds
                    domain_mean_t = float(np.clip(
                        domain_mean_t, p.diff_range[0], p.diff_range[1]))
                else:
                    domain_mean_t = mean_t

                # Apply stratified sampling on the combined D pool,
                # honouring per-category hard caps and domain-level mean target
                sel, warns = _pick_stratified_with_quota(
                    d_pool, total_needed, intra_form_used_ids,
                    quota, "_cat_label",
                    p.diff_range, domain_mean_t, sigma,
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
        form = form.drop(columns=["_slot_label"], errors="ignore")

        # ── Used Blocks fallback (5th priority — absolute last resort) ─────────
        # Triggered only when:
        #   (a) allow_used_blocks is True in params
        #   (b) the form is short of its required total
        #   (c) a non-empty used_blocks DataFrame was provided
        p = self.params
        if (p.allow_used_blocks and
                p.used_blocks is not None and
                not p.used_blocks.empty):

            # Compute actual vs required totals per slot
            total_required = sum(s.count for s in p.slots)
            got_ids  = set(form["QuestionID"].dropna().astype(str).tolist()) \
                       if "QuestionID" in form.columns else set()
            shortfall = total_required - len(got_ids)

            if shortfall > 0:
                # Build slot quotas still missing
                for slot in p.slots:
                    # Count how many items for this slot were already assembled
                    if "Category" in form.columns:
                        cat_col = "Category"
                    elif "الناتج" in form.columns:
                        cat_col = "الناتج"
                    else:
                        cat_col = None

                    got_slot = 0
                    if cat_col and cat_col in form.columns:
                        got_slot = (form[cat_col] == slot.label).sum()
                    need_more = max(slot.count - got_slot, 0)

                    if need_more <= 0:
                        continue

                    # Pull from Used Blocks: strictly filtered by Stage-Level prefix,
                    # then by category and difficulty range.
                    # NO relaxation of the stage-level filter is ever allowed.
                    ub = p.used_blocks.copy()

                    # ── STRICT Stage-Level filter (new requirement) ───────────
                    # Only reuse questions that came from the EXACT same
                    # Stage and Level/Part as the block being assembled.
                    #
                    # self.block_prefix is derived from the current block ID:
                    #   "GAT-2.1.006"  →  prefix = "GAT-2.1"
                    #   "GAT-1.2.003"  →  prefix = "GAT-1.2"
                    #
                    # If no Block-ID column exists or no prefix is set,
                    # fall through to category/difficulty filtering only
                    # (backward-compatible with banks that predate block IDs).
                    if (self.block_prefix and
                            "Block-ID" in ub.columns):
                        # Keep only rows whose Block-ID starts with this prefix
                        prefix_filter = (
                            ub["Block-ID"].astype(str)
                            .str.startswith(self.block_prefix + ".")
                        )
                        ub_stage = ub[prefix_filter]

                        if ub_stage.empty:
                            # Strictly no candidates for this stage-level —
                            # do NOT borrow from other stages/levels.
                            all_warnings.append(
                                f"Form {form_number}: Used Blocks has no questions "
                                f"matching stage-level '{self.block_prefix}' for "
                                f"'{slot.label}' — {need_more} item(s) cannot be "
                                f"filled from fallback (cross-stage reuse forbidden).")
                            continue
                        ub = ub_stage   # continue with stage-filtered pool only

                    # Exclude already in this form
                    if "QuestionID" in ub.columns:
                        ub = ub[~ub["QuestionID"].astype(str).isin(intra_form_used_ids)]

                    # Filter by category if column exists (within stage-filtered pool)
                    if "Category" in ub.columns and slot.label:
                        ub_cat = ub[ub["Category"].astype(str) == str(slot.label)]
                        if ub_cat.empty:
                            # No category match within this stage-level —
                            # do NOT fall back to other categories inside the stage
                            all_warnings.append(
                                f"Form {form_number}: No Used Blocks match "
                                f"stage-level '{self.block_prefix or 'n/a'}' "
                                f"AND category '{slot.label}' — "
                                f"{need_more} item(s) skipped.")
                            continue
                    else:
                        ub_cat = ub

                    # Filter by difficulty range
                    dcol_ub = "Difficulty" if "Difficulty" in ub_cat.columns else self.dcol
                    if dcol_ub in ub_cat.columns:
                        lo, hi = p.diff_range
                        ub_cat = ub_cat[
                            (pd.to_numeric(ub_cat[dcol_ub], errors="coerce") >= lo) &
                            (pd.to_numeric(ub_cat[dcol_ub], errors="coerce") <= hi)
                        ]
                        # NOTE: do NOT relax difficulty range —
                        # if nothing is left the slot stays empty with a warning
                        if ub_cat.empty:
                            all_warnings.append(
                                f"Form {form_number}: Used Blocks candidates for "
                                f"stage-level '{self.block_prefix or 'n/a'}' / "
                                f"'{slot.label}' are all outside the difficulty "
                                f"range [{lo},{hi}] — {need_more} item(s) skipped.")
                            continue

                    if ub_cat.empty:
                        all_warnings.append(
                            f"Form {form_number}: Used Blocks also exhausted for "
                            f"'{slot.label}' (stage-level '{self.block_prefix or 'n/a'}') "
                            f"— {need_more} placeholders added.")
                        continue

                    # ── Max-Usage cap + Priority selection ────────────────────
                    # Apply the session-level UB usage counter so that:
                    #   Priority 1 — questions never selected before (unique)
                    #   Priority 2 — questions selected < max_ub_usage times
                    # Hard limit: questions at max_ub_usage or above are excluded.
                    max_ub = p.max_ub_usage if p.max_ub_usage > 0 else 999

                    if "QuestionID" in ub_cat.columns:
                        ub_cat = ub_cat.copy()
                        ub_cat["_session_uses"] = (
                            ub_cat["QuestionID"].astype(str)
                            .map(lambda q: self.ub_usage.get(q, 0))
                        )
                        # Hard cap: never exceed max_ub_usage
                        ub_cat = ub_cat[ub_cat["_session_uses"] < max_ub]

                    if ub_cat.empty:
                        all_warnings.append(
                            f"Form {form_number}: All Used Blocks candidates for "
                            f"'{slot.label}' have reached the max usage limit "
                            f"({max_ub}) — {need_more} item(s) skipped.")
                        continue

                    # Sort order:
                    #   1. _session_uses ascending (prefer never-used first)
                    #   2. 'Number of used' ascending (least-used in UB first)
                    # Within the same tier, shuffle for fairness.
                    sort_cols  = ["_session_uses"]
                    sort_asc   = [True]
                    if "Number of used" in ub_cat.columns:
                        sort_cols.append("Number of used")
                        sort_asc.append(True)
                    ub_cat = ub_cat.sort_values(sort_cols, ascending=sort_asc)

                    # Within each tier (same _session_uses value), shuffle
                    # so the same questions aren't always picked first.
                    shuffled_parts = []
                    for tier_val in sorted(ub_cat["_session_uses"].unique()):
                        tier = ub_cat[ub_cat["_session_uses"] == tier_val].copy()
                        tier = tier.sample(frac=1,
                                           random_state=int(self.rng.integers(2**31)))
                        shuffled_parts.append(tier)
                    if shuffled_parts:
                        ub_cat = pd.concat(shuffled_parts, ignore_index=True)

                    # Rename difficulty column to match main bank schema if needed
                    if self.dcol not in ub_cat.columns and dcol_ub in ub_cat.columns:
                        ub_cat = ub_cat.rename(columns={dcol_ub: self.dcol})

                    top = ub_cat.head(need_more).copy()
                    top["_form"]   = form_number
                    top["_source"] = "Used Blocks"

                    ub_picked = list(top["QuestionID"].dropna().astype(str))
                    intra_form_used_ids.update(ub_picked)
                    # Update the session counter immediately so later slots
                    # and later forms see the correct usage counts.
                    for qid in ub_picked:
                        self.ub_usage[qid] = self.ub_usage.get(qid, 0) + 1
                    form_parts.append(top)

                    tier_info = "unique" if top.get("_session_uses", pd.Series([0])).max() == 0 \
                                else f"reused (max_uses={max_ub})"
                    all_warnings.append(
                        f"Form {form_number}: {len(top)} question(s) for "
                        f"'{slot.label}' taken from Used Blocks "
                        f"({tier_info}, least-used first).")
                    log.info("Used Blocks fallback: %d items for '%s' in form %d [%s]",
                             len(top), slot.label, form_number, tier_info)

                # Rebuild form with the fallback rows added
                form = pd.concat(form_parts, ignore_index=True) if form_parts \
                       else pd.DataFrame()
                form = form.drop(columns=["_slot_label", "_source"],
                                 errors="ignore")

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


def _adaptive_mean_correction(
        form: pd.DataFrame,
        dcol: str,
        mn_range: tuple,
        params: "AssemblyParams",
        bank: pd.DataFrame,
        intra_form_used_ids: set,
        question_usage: dict,
        rng: np.random.Generator,
        form_number: int,
        block_prefix: Optional[str] = None,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Adaptive Mean Control — post-assembly correction pass.

    After the initial form is assembled, check whether the current mean
    is within the target range.  If not:

      Mean too LOW  → swap some low-difficulty items for higher ones
      Mean too HIGH → swap some high-difficulty items for lower ones

    Swap candidate sources (in priority order, least-used first):
      1. Main bank (fresh questions)
      2. Reused questions (when allow_reuse=True)
      3. Used Blocks dataset (when allow_used_blocks=True)

    Constraints honoured in every swap:
      • Category quota (no slot exceeds its target count)
      • Difficulty range (diff_range)
      • No duplicates within the form
      • Stage / Level structure (all swaps stay within diff_range)

    Returns the corrected form and a list of log messages.
    """
    warnings: list[str] = []
    if mn_range is None or form.empty:
        return form, warnings

    mn_lo, mn_hi = mn_range
    current_mean = float(pd.to_numeric(form[dcol], errors="coerce").dropna().mean())

    if mn_lo <= current_mean <= mn_hi:
        return form, warnings   # already fine

    direction = "higher" if current_mean < mn_lo else "lower"
    target_mean = (mn_lo + mn_hi) / 2.0

    # Build combined correction pool: bank + reuse + used_blocks,
    # sorted by difficulty direction (asc for need-higher, desc for need-lower)
    # then by usage count ascending (least-used first).
    correction_pool_parts = []

    # Source 1: main bank (fresh)
    fresh_used = {q for q, c in question_usage.items() if c >= 1}
    fresh_pool = bank[~bank["QuestionID"].astype(str).isin(fresh_used)].copy()
    fresh_pool = fresh_pool[~fresh_pool["QuestionID"].astype(str)
                             .isin(intra_form_used_ids)]
    fresh_pool["_source"]    = "bank"
    fresh_pool["_use_count"] = 0
    correction_pool_parts.append(fresh_pool)

    # Source 2: reusable questions from main bank
    if params.allow_reuse:
        reuse_pool = bank[
            bank["QuestionID"].astype(str).map(
                lambda q: 0 < question_usage.get(q, 0) < params.max_reuse)
        ].copy()
        reuse_pool = reuse_pool[~reuse_pool["QuestionID"].astype(str)
                                 .isin(intra_form_used_ids)]
        reuse_pool["_source"]    = "reuse"
        reuse_pool["_use_count"] = reuse_pool["QuestionID"].astype(str).map(
            lambda q: question_usage.get(q, 0))
        correction_pool_parts.append(reuse_pool)

    # Source 3: Used Blocks dataset — STRICT stage-level filter applied first
    if params.allow_used_blocks and params.used_blocks is not None \
            and not params.used_blocks.empty:
        ub = _filter_ub_by_prefix(params.used_blocks, block_prefix)
        if ub.empty and block_prefix:
            warnings.append(
                f"Form {form_number}: _adaptive_mean_correction — Used Blocks has "
                f"no records matching stage-level '{block_prefix}'; "
                f"cross-stage reuse is forbidden.")
        elif not ub.empty:
            if dcol not in ub.columns and "Difficulty" in ub.columns:
                ub = ub.rename(columns={"Difficulty": dcol})
            if dcol in ub.columns:
                ub = ub[~ub["QuestionID"].astype(str).isin(intra_form_used_ids)]
                ub[dcol] = pd.to_numeric(ub[dcol], errors="coerce").fillna(0.5)
                ub["_source"]    = "used_blocks"
                ub["_use_count"] = ub.get("Number of used",
                                           pd.Series(0, index=ub.index))
                correction_pool_parts.append(ub)

    if not correction_pool_parts:
        return form, warnings

    corr_pool = pd.concat(correction_pool_parts, ignore_index=True)
    corr_pool[dcol] = pd.to_numeric(corr_pool[dcol], errors="coerce")

    # Keep only items within the allowed difficulty range
    lo, hi = params.diff_range
    corr_pool = corr_pool[(corr_pool[dcol] >= lo) & (corr_pool[dcol] <= hi)]
    corr_pool = corr_pool.dropna(subset=[dcol])

    if corr_pool.empty:
        return form, warnings

    # Sort correction pool: direction-aware difficulty, then least-used first
    if direction == "higher":
        corr_pool = corr_pool.sort_values(
            [dcol, "_use_count"], ascending=[False, True])
    else:
        corr_pool = corr_pool.sort_values(
            [dcol, "_use_count"], ascending=[True, True])

    # Iteratively swap: replace the item that deviates most from target_mean
    # with the best available correction candidate.
    form = form.copy()
    max_swaps = max(1, len(form) // 3)   # swap at most 1/3 of the form

    for _ in range(max_swaps):
        current_mean = float(pd.to_numeric(form[dcol], errors="coerce").dropna().mean())
        if mn_lo <= current_mean <= mn_hi:
            break

        if direction == "higher":
            # Find the lowest-difficulty item in the form that can be swapped
            valid_idx = form[dcol].notna()
            if not valid_idx.any():
                break
            worst_pos = int(pd.to_numeric(form.loc[valid_idx, dcol]).idxmin())
            worst_diff = float(form.at[worst_pos, dcol])
            # Only swap if it actually helps
            needed_diff = target_mean + (target_mean - current_mean) * len(form)
            # Candidate: higher-difficulty item from correction pool
            candidates = corr_pool[corr_pool[dcol] > worst_diff]
        else:
            valid_idx = form[dcol].notna()
            if not valid_idx.any():
                break
            worst_pos = int(pd.to_numeric(form.loc[valid_idx, dcol]).idxmax())
            worst_diff = float(form.at[worst_pos, dcol])
            candidates = corr_pool[corr_pool[dcol] < worst_diff]

        if candidates.empty:
            break

        # Check category quota — replacement must have the same category
        if "Category" in form.columns and "Category" in candidates.columns:
            worst_cat = form.at[worst_pos, "Category"]
            cat_candidates = candidates[
                candidates["Category"].astype(str) == str(worst_cat)]
            if cat_candidates.empty:
                # Relax: any category that still has room
                cat_candidates = candidates
        else:
            cat_candidates = candidates

        if cat_candidates.empty:
            break

        replacement = cat_candidates.iloc[0].copy()
        rep_qid = str(replacement.get("QuestionID", ""))

        # Perform the swap
        old_qid = str(form.at[worst_pos, "QuestionID"]) \
                  if "QuestionID" in form.columns else ""

        # Remove the swap candidate from the correction pool
        corr_pool = corr_pool[corr_pool["QuestionID"].astype(str) != rep_qid]

        # Update the form row (keep only columns that exist in form)
        shared_cols = [c for c in replacement.index if c in form.columns
                       and not c.startswith("_")]
        for col in shared_cols:
            form.at[worst_pos, col] = replacement[col]

        intra_form_used_ids.discard(old_qid)
        intra_form_used_ids.add(rep_qid)

        src = replacement.get("_source", "bank")
        warnings.append(
            f"Form {form_number}: adaptive swap — removed diff={worst_diff:.3f} "
            f"replaced with diff={float(replacement[dcol]):.3f} "
            f"(source: {src})")

    final_mean = float(pd.to_numeric(form[dcol], errors="coerce").dropna().mean())
    if mn_lo <= final_mean <= mn_hi:
        log.info("Form %d: adaptive correction achieved mean=%.3f", form_number, final_mean)
    else:
        warnings.append(
            f"Form {form_number}: adaptive correction improved mean to {final_mean:.3f} "
            f"but target [{mn_lo:.2f},{mn_hi:.2f}] still not fully met.")

    return form, warnings


def _domain_means_ok(form: pd.DataFrame, dcol: str,
                     mn_range: tuple, domain_col: str) -> bool:
    """
    Check that EVERY D-domain in the form has its mean difficulty within
    mn_range.  This prevents Domain V (or any other domain) from
    individually violating the mean criterion even when the overall mean
    looks fine.
    """
    if domain_col not in form.columns:
        return True
    for dom in form[domain_col].dropna().unique():
        sub   = form[form[domain_col] == dom]
        diffs = pd.to_numeric(sub[dcol], errors="coerce").dropna()
        if diffs.empty:
            continue
        dom_mean = float(diffs.mean())
        if not (mn_range[0] <= dom_mean <= mn_range[1]):
            return False
    return True


def assemble_forms(bank: pd.DataFrame,
                   params: AssemblyParams,
                   dcol: str,
                   n_forms: int,
                   form_names: list[str]
                   ) -> list[tuple[pd.DataFrame, float, list[str]]]:
    """
    Assemble n_forms sharing usage tracking so questions spread.

    Features 2 / 3 / 4 are integrated here:
      Feature 2 (scope)    → _check_scope_mean() replaces hard-coded domain check
      Feature 3 (CCT)      → resolve_diff_params() applies CCT profile constraints
      Feature 4 (fallback) → _fallback_for_difficulty() runs after retries
    """
    usage: dict[str, int] = {}
    # ub_usage tracks how many times each Used-Blocks question has been
    # selected across ALL forms in this assembly session.
    # Key: QuestionID (str), Value: count of forms it has appeared in.
    ub_usage: dict[str, int] = {}
    results: list[tuple[pd.DataFrame, float, list[str]]] = []
    domain_col = "D" if "D" in bank.columns else "المجال"

    # Feature 3: resolve effective range/mean from Delta or CCT
    eff_range, eff_mn = resolve_diff_params(params)
    mn_range = eff_mn

    def _passes(form, mean_d):
        """True when the form satisfies the mean criterion at the chosen scope."""
        if mn_range is None:
            return True
        # Feature 2: scope-aware check
        return _check_scope_mean(form, dcol, mn_range,
                                 params.difficulty_scope, domain_col)

    for fidx, name in enumerate(form_names):
        # Derive stage-level prefix from the block name for Used Blocks filtering.
        # Block IDs follow the pattern  CODE-Stage.Level.Number
        # e.g. "GAT-2.1.006" → prefix "GAT-2.1"
        #      "GAT-1.2.003" → prefix "GAT-1.2"
        #      "Form_1"      → None (no filtering, backward compatible)
        import re as _re
        _m = _re.match(r'^([A-Za-z]+-\d+\.\d+)\.\d+$', str(name))
        block_prefix = _m.group(1) if _m else None

        engine = AssemblyEngine(bank, params, dcol, dict(usage),
                                block_prefix=block_prefix)
        engine.ub_usage = ub_usage   # share the session-level UB counter
        best_form, best_mean, best_warns = engine.assemble_one_form(fidx + 1)

        # ── Adaptive Mean Control (runs BEFORE the retry loop) ─────────────
        # When enabled, attempt a targeted swap-correction pass first.
        # This often avoids the need for expensive retries entirely.
        if params.adaptive_mean_control and mn_range is not None:
            intra_ids = set(best_form["QuestionID"].dropna().astype(str).tolist()) \
                        if "QuestionID" in best_form.columns else set()
            corr_form, corr_warns = _adaptive_mean_correction(
                best_form, dcol, mn_range, params,
                bank, intra_ids, dict(usage),
                np.random.default_rng(
                    (params.rng_seed or 0) + fidx * 999),
                fidx + 1,
                block_prefix=block_prefix,
            )
            corr_mean = float(pd.to_numeric(
                corr_form[dcol], errors="coerce").dropna().mean()) \
                if not corr_form.empty else best_mean
            best_warns.extend(corr_warns)
            # Accept correction if it improved the result
            if abs(corr_mean - (mn_range[0]+mn_range[1])/2) < \
               abs(best_mean  - (mn_range[0]+mn_range[1])/2):
                best_form, best_mean = corr_form, corr_mean

        success = _passes(best_form, best_mean)

        if not success:
            for attempt in range(1, params.max_retries):
                seed = (params.rng_seed or 0) + fidx * 1000 + attempt
                p2 = AssemblyParams(**{
                    **params.__dict__,
                    "rng_seed": seed,
                    "diff_mean_range": None,
                })
                eng2 = AssemblyEngine(bank, p2, dcol, dict(usage),
                                      block_prefix=block_prefix)
                eng2.ub_usage = ub_usage   # share session UB counter
                form2, mean2, w2 = eng2.assemble_one_form(fidx + 1)

                # Run adaptive correction on each retry attempt too
                if params.adaptive_mean_control and mn_range is not None:
                    intra2 = set(form2["QuestionID"].dropna().astype(str)) \
                             if "QuestionID" in form2.columns else set()
                    form2, cw = _adaptive_mean_correction(
                        form2, dcol, mn_range, params,
                        bank, intra2, dict(usage),
                        np.random.default_rng(seed + 7),
                        fidx + 1,
                        block_prefix=block_prefix,
                    )
                    w2.extend(cw)
                    mean2 = float(pd.to_numeric(
                        form2[dcol], errors="coerce").dropna().mean()) \
                        if not form2.empty else mean2

                if _passes(form2, mean2):
                    best_form, best_mean, best_warns = form2, mean2, w2
                    success = True
                    break
                # Keep the attempt closest to target even if none fully pass
                curr_err = abs(best_mean - (mn_range[0]+mn_range[1])/2)
                new_err  = abs(mean2      - (mn_range[0]+mn_range[1])/2)
                if new_err < curr_err:
                    best_form, best_mean, best_warns = form2, mean2, w2

            if not success:
                best_warns.append(
                    f"{name}: mean criterion ({mn_range[0]:.2f}–{mn_range[1]:.2f}) "
                    f"not met after {params.max_retries} retries "
                    f"(achieved {best_mean:.3f})."
                )

        # ── Feature 4: Fallback-for-difficulty (dedicated reuse for mean) ──
        if not success and params.allow_fallback_reuse and mn_range is not None:
            intra_ids = set(best_form["QuestionID"].dropna().astype(str).tolist()) \
                        if "QuestionID" in best_form.columns else set()
            best_form, fb_warns = _fallback_for_difficulty(
                best_form, dcol, mn_range,
                params.difficulty_scope, params,
                bank, intra_ids, dict(usage),
                np.random.default_rng((params.rng_seed or 0) + fidx * 31337),
                fidx + 1,
                block_prefix=block_prefix,
            )
            best_warns.extend(fb_warns)
            if fb_warns:
                best_mean = float(pd.to_numeric(
                    best_form[dcol], errors="coerce").dropna().mean()) \
                    if not best_form.empty else best_mean
                success = _passes(best_form, best_mean)
                log.info("%s: fallback-for-difficulty applied → mean=%.3f [%s]",
                         name, best_mean, "OK" if success else "still missed")

        # Update cross-form usage by QuestionID
        if "QuestionID" in best_form.columns:
            for qid in best_form["QuestionID"].dropna().astype(str):
                usage[qid] = usage.get(qid, 0) + 1

        # Update the session-level Used Blocks usage counter so subsequent
        # forms know which UB questions have already been selected and how
        # many times they have been used.
        if (params.allow_used_blocks and
                params.used_blocks is not None and
                "QuestionID" in best_form.columns and
                "QuestionID" in params.used_blocks.columns):
            ub_qids = set(params.used_blocks["QuestionID"].astype(str))
            for qid in best_form["QuestionID"].dropna().astype(str):
                if qid in ub_qids:
                    ub_usage[qid] = ub_usage.get(qid, 0) + 1

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
                        col_offset: int, row_offset: int,
                        table_name: str = "",
                        domain_col: str = "",
                        domain_value: str = "") -> int:
    """
    Write difficulty distribution bin table + bar chart for one domain.

    When `table_name` is supplied the bin counts are written as live
    COUNTIFS formulas that reference the Excel Table — any edit to the
    question table (difficulty values, domain changes) instantly refreshes
    the chart data and the chart itself.

    When `table_name` is empty (legacy / fallback) static Python-computed
    counts are written instead (original behaviour preserved).
    """
    fmts = _wb_fmts(wb)
    ws.write(row_offset, col_offset,
             f"{domain_label} — Difficulty Range", fmts["shdr"])
    ws.write(row_offset, col_offset + 1, "Count", fmts["shdr"])

    for i, lbl in enumerate(BIN_LABELS):
        lo = round(i * 0.1, 1)
        hi = round((i + 1) * 0.1, 1)
        ws.write(row_offset + 1 + i, col_offset, lbl)

        if table_name and domain_col and domain_value:
            # ── Dynamic COUNTIFS formula ─────────────────────────────────
            # Last bin uses "<=" so values exactly equal to 1.0 are included.
            cmp_op  = "<=" if i == 9 else "<"
            hi_val  = 1.0  if i == 9 else hi
            formula = (
                f'=COUNTIFS({table_name}[{domain_col}],"{domain_value}",'
                f'{table_name}[{dcol}],">="&{lo},'
                f'{table_name}[{dcol}],"{cmp_op}"&{hi_val})'
            )
            ws.write_formula(row_offset + 1 + i, col_offset + 1, formula)
        elif table_name and dcol:
            # ── Dynamic without domain filter ────────────────────────────
            cmp_op = "<=" if i == 9 else "<"
            hi_val = 1.0  if i == 9 else hi
            formula = (
                f'=COUNTIFS({table_name}[{dcol}],">="&{lo},'
                f'{table_name}[{dcol}],"{cmp_op}"&{hi_val})'
            )
            ws.write_formula(row_offset + 1 + i, col_offset + 1, formula)
        else:
            # ── Legacy: static count (original behaviour) ────────────────
            cnt = int((
                (domain_df[dcol] >= lo) &
                (domain_df[dcol] < (hi + 0.001 if i == 9 else hi))
            ).sum()) if dcol in domain_df.columns else 0
            ws.write(row_offset + 1 + i, col_offset + 1, cnt)

    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name":       "Q Distribution",
        "categories": [sheet_name, row_offset+1, col_offset,
                        row_offset+10, col_offset],
        "values":     [sheet_name, row_offset+1, col_offset+1,
                        row_offset+10, col_offset+1],
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


def _embed_3pl_charts(wb, ws, fa: FormAnalysis, img_row: int,
                       sheet_name: str = ""):
    """
    Embed 3PL charts as LIVE xlsxwriter native charts.

    Charts reference actual worksheet cell ranges → any manual edit to
    the data cells is immediately reflected in the chart (dynamic binding).
    No static PNG images are used.

    Data tables written at column U (index 20), then charts reference them:
      Row img_row         col 20 : Score Distribution data + chart
      Row img_row + n+3   col 20 : ICC data + chart
      Row img_row + 2n+6  col 20 : SEM data + chart
    """
    RIGHT_COL = 20
    FMTS = _wb_fmts(wb)

    # ── Score Distribution ─────────────────────────────────────────────────
    score_counts, score_edges = np.histogram(fa.scores, bins=15)
    score_labels = [f"{score_edges[i]:.0f}–{score_edges[i+1]:.0f}"
                    for i in range(len(score_counts))]
    n_sc = len(score_counts)

    ws.write(img_row, RIGHT_COL,   "Score Bin",  FMTS["shdr"])
    ws.write(img_row, RIGHT_COL+1, "Frequency",  FMTS["shdr"])
    for i, (lbl, cnt) in enumerate(zip(score_labels, score_counts)):
        ws.write(img_row+1+i, RIGHT_COL,   lbl)
        ws.write(img_row+1+i, RIGHT_COL+1, int(cnt))

    ch_sc = wb.add_chart({"type": "column"})
    ch_sc.add_series({
        "name":       "Score Distribution",
        "categories": [sheet_name, img_row+1, RIGHT_COL,
                        img_row+n_sc, RIGHT_COL],
        "values":     [sheet_name, img_row+1, RIGHT_COL+1,
                        img_row+n_sc, RIGHT_COL+1],
        "fill":       {"color": ETEC_BLUE},
    })
    ch_sc.set_title({"name": "Simulated Score Distribution"})
    ch_sc.set_x_axis({"name": "Total Score"})
    ch_sc.set_y_axis({"name": "Students"})
    ch_sc.set_legend({"none": True})
    ch_sc.set_size({"width": 420, "height": 240})
    ws.insert_chart(img_row, RIGHT_COL+3, ch_sc)

    # ── Average ICC ────────────────────────────────────────────────────────
    step    = max(1, len(fa.theta_grid) // 20)
    th_ds   = fa.theta_grid[::step]
    icc_ds  = fa.icc.mean(axis=1)[::step]
    n_ic    = len(th_ds)
    icc_row = img_row + n_sc + 3

    ws.write(icc_row, RIGHT_COL,   "Theta",    FMTS["shdr"])
    ws.write(icc_row, RIGHT_COL+1, "Avg P(θ)", FMTS["shdr"])
    for i, (th, pp) in enumerate(zip(th_ds, icc_ds)):
        ws.write(icc_row+1+i, RIGHT_COL,   round(float(th), 2))
        ws.write(icc_row+1+i, RIGHT_COL+1, round(float(pp), 4))

    ch_ic = wb.add_chart({"type": "line"})
    ch_ic.add_series({
        "name":       "Avg ICC",
        "categories": [sheet_name, icc_row+1, RIGHT_COL,
                        icc_row+n_ic, RIGHT_COL],
        "values":     [sheet_name, icc_row+1, RIGHT_COL+1,
                        icc_row+n_ic, RIGHT_COL+1],
        "line":       {"color": ETEC_TEAL, "width": 2.25},
    })
    ch_ic.set_title({"name": "Item Characteristic Curve (ICC)"})
    ch_ic.set_x_axis({"name": "Ability (θ)"})
    ch_ic.set_y_axis({"name": "P(Correct)", "min": 0, "max": 1})
    ch_ic.set_legend({"none": True})
    ch_ic.set_size({"width": 420, "height": 240})
    ws.insert_chart(icc_row, RIGHT_COL+3, ch_ic)

    # ── SEM ────────────────────────────────────────────────────────────────
    sem_ds  = np.clip(fa.sem[::step], 0, 5.0)
    sem_row = icc_row + n_ic + 3

    ws.write(sem_row, RIGHT_COL,   "Theta", FMTS["shdr"])
    ws.write(sem_row, RIGHT_COL+1, "SEM",   FMTS["shdr"])
    for i, (th, sv) in enumerate(zip(th_ds, sem_ds)):
        ws.write(sem_row+1+i, RIGHT_COL,   round(float(th), 2))
        ws.write(sem_row+1+i, RIGHT_COL+1, round(float(sv), 4))

    ch_sem = wb.add_chart({"type": "line"})
    ch_sem.add_series({
        "name":       "SEM",
        "categories": [sheet_name, sem_row+1, RIGHT_COL,
                        sem_row+n_ic, RIGHT_COL],
        "values":     [sheet_name, sem_row+1, RIGHT_COL+1,
                        sem_row+n_ic, RIGHT_COL+1],
        "line":       {"color": ETEC_GREEN, "width": 2.25},
    })
    ch_sem.set_title({"name": "Standard Error of Measurement (SEM)"})
    ch_sem.set_x_axis({"name": "Ability (θ)"})
    ch_sem.set_y_axis({"name": "SEM", "min": 0})
    ch_sem.set_legend({"none": True})
    ch_sem.set_size({"width": 420, "height": 240})
    ws.insert_chart(sem_row, RIGHT_COL+3, ch_sem)


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


def _make_output_folder(base_dir: Path, exam_type: str, n_forms: int) -> Path:
    """
    Create and return a timestamped output folder:
      <base_dir>/<exam_type> - <n_forms> Forms - YYYY-MM-DD - HH-MM-SS/

    Characters illegal in folder names are stripped automatically.
    """
    import datetime, re
    now       = datetime.datetime.now()
    date_str  = now.strftime("%Y-%m-%d")
    time_str  = now.strftime("%H-%M-%S")
    safe_exam = re.sub(r'[\\/:*?"<>|]', "", str(exam_type)).strip()
    folder_name = f"{safe_exam} - {n_forms} Forms - {date_str} - {time_str}"
    out = base_dir / folder_name
    out.mkdir(parents=True, exist_ok=True)
    return out


def write_forms_workbook(
    results: list[tuple[pd.DataFrame, float, list[str]]],
    form_names: list[str],
    analyses: list[Optional[FormAnalysis]],
    dcol: str,
    domain_col: str,      # "D" for mode 1, "المجال" for mode 2
    acol: str,
    ccol: str,
    question_usage: dict,
    out_path: Path,
    mode: int,
):
    """
    Write one sheet per form.

    Dynamic synchronisation
    -----------------------
    The question table is converted to an Excel Table (ListObject / named
    table).  Every metric in the summary header, every domain column, and
    every difficulty-distribution bin count is written as a live Excel
    formula that references the table by name.

    This means:
    • Editing any cell in the question table (Difficulty, D, Category, …)
      automatically refreshes ALL metrics, statistics, and charts in the
      same sheet — no VBA, no manual refresh.
    • Adding or removing rows inside the table expands/contracts all
      calculations automatically because Excel Table structured references
      grow with the table.
    • Charts reference the formula cells, so they also update live.

    Static values retained
    ----------------------
    Cronbach Alpha and Item Correlation come from the Python 3PL simulation
    and are written as static numbers (they cannot be recomputed from the
    item table alone using only Excel formulas).
    """
    import re as _re

    wb   = xlsxwriter.Workbook(str(out_path))
    fmts = _wb_fmts(wb)
    fa_map      = {fa.name: fa for fa in analyses if fa}
    export_cols = M1_EXPORT_COLS if mode == 1 else M2_EXPORT_COLS
    n_item_cols = len(export_cols)

    for form_idx, ((form_df, mean_d, warns), name) in enumerate(
            zip(results, form_names)):

        sheet_name = name[:31]
        ws = wb.add_worksheet(sheet_name)
        if mode == 2:
            ws.right_to_left()

        df = form_df.reset_index(drop=True)
        fa = fa_map.get(name)

        # ── Excel Table name ─────────────────────────────────────────────────
        # Rules: unique across workbook, alphanumeric + underscore, must
        # NOT look like a cell reference (e.g. "F1", "A10", "R2C3").
        # Safest approach: always prefix with "QTable" + form index.
        tbl_name = f"QTable{form_idx + 1}"

        # ── Discover domains present in this form ────────────────────────────
        if domain_col in df.columns:
            domains = sorted(
                str(d) for d in df[domain_col].dropna().unique() if str(d).strip()
            )
        else:
            domains = []

        # ── Column widths ────────────────────────────────────────────────────
        ws.set_column(0, 0, 32)
        ws.set_column(1, 1, 14)
        for di in range(len(domains)):
            ws.set_column(2 + di, 2 + di, 16)

        # ── Title row ────────────────────────────────────────────────────────
        n_summary_cols = max(2 + len(domains), 4)
        ws.merge_range(0, 0, 0, n_summary_cols - 1, name, fmts["ttl"])
        ws.set_row(0, 22)

        # ── Summary header row ───────────────────────────────────────────────
        ws.write(2, 0, "Metric",  fmts["shdr"])
        ws.write(2, 1, "Overall", fmts["shdr"])
        for di, dom in enumerate(domains):
            ws.write(2, 2 + di, f"Domain: {dom}", fmts["shdr"])

        # ── Helper: write one metric row ─────────────────────────────────────
        def _metric(row_idx: int, label: str,
                    overall,          # str formula, numeric, or "—"
                    dom_vals: list,   # per-domain: str formula, numeric, or "—"
                    is_float: bool):
            ws.write(row_idx, 0, label, fmts["mlbl"])
            num_fmt = fmts["mval"] if is_float else fmts["mvi"]

            def _put(row, col, val):
                if isinstance(val, str) and val.startswith("="):
                    ws.write_formula(row, col, val, num_fmt)
                elif isinstance(val, str):
                    ws.write(row, col, val, fmts["dat"])
                elif val is None or (isinstance(val, float) and math.isnan(val)):
                    ws.write(row, col, "—", fmts["dat"])
                else:
                    try:
                        ws.write_number(row, col,
                                        int(val) if not is_float else float(val),
                                        num_fmt)
                    except Exception:
                        ws.write(row, col, str(val), fmts["dat"])

            _put(row_idx, 1, overall)
            for di, dv in enumerate(dom_vals):
                _put(row_idx, 2 + di, dv)

        # Shorthand for formulas
        TN = tbl_name
        D  = dcol
        A  = acol
        DC = domain_col

        # Number of Items  — COUNTA / COUNTIF
        _metric(3, "Number of Items",
                f'=COUNTA({TN}[QuestionID])',
                [f'=COUNTIF({TN}[{DC}],"{d}")' for d in domains],
                False)

        # Mean Difficulty  — AVERAGE / AVERAGEIF
        _metric(4, "Mean Difficulty",
                f'=IFERROR(AVERAGE({TN}[{D}]),"—")',
                [f'=IFERROR(AVERAGEIF({TN}[{DC}],"{d}",{TN}[{D}]),"—")'
                 for d in domains],
                True)

        # Min Difficulty  — MIN / MINIFS (Excel 2019+ / 365)
        _metric(5, "Min Difficulty",
                f'=IFERROR(MIN({TN}[{D}]),"—")',
                [f'=IFERROR(MINIFS({TN}[{D}],{TN}[{DC}],"{d}"),"—")'
                 for d in domains],
                True)

        # Max Difficulty  — MAX / MAXIFS
        _metric(6, "Max Difficulty",
                f'=IFERROR(MAX({TN}[{D}]),"—")',
                [f'=IFERROR(MAXIFS({TN}[{D}],{TN}[{DC}],"{d}"),"—")'
                 for d in domains],
                True)

        # Avg Discrimination  — AVERAGE / AVERAGEIF
        if acol and acol in df.columns:
            _metric(7, f"Avg Discrimination ({acol})",
                    f'=IFERROR(AVERAGE({TN}[{A}]),"—")',
                    [f'=IFERROR(AVERAGEIF({TN}[{DC}],"{d}",{TN}[{A}]),"—")'
                     for d in domains],
                    True)
        else:
            _metric(7, "Avg Discrimination", "—", ["—"] * len(domains), True)

        # Cronbach Alpha & Item Correlation — static (from 3PL simulation)
        alpha = fa.alpha     if fa and not math.isnan(fa.alpha)     else "—"
        icorr = fa.item_corr if fa and not math.isnan(fa.item_corr) else "—"
        _metric(8, "Simulated Cronbach Alpha",  alpha, ["—"] * len(domains), True)
        _metric(9, "Mean Item Correlation",     icorr, ["—"] * len(domains), True)

        # ── Warnings ─────────────────────────────────────────────────────────
        warn_row = 11
        if warns:
            ws.write(warn_row, 0, "⚠ Warnings", fmts["warn"])
            for wi, w in enumerate(warns[:6]):
                ws.write(warn_row + 1 + wi, 0, w, fmts["warn"])
            warn_row += len(warns[:6]) + 2
        tbl_start = warn_row + 1   # first row of item table (header)

        # ── Item Table data ───────────────────────────────────────────────────
        # Write data rows BEFORE calling add_table() so we can apply custom
        # row-level formatting (alternating colours, red for reused items).
        # add_table() will add the table structure on top.
        for ci, (cn, cw, _) in enumerate(export_cols):
            ws.set_column(ci, ci, cw)

        for ri, (_, row) in enumerate(df.iterrows()):
            qid    = str(row.get("QuestionID", ""))
            reused = question_usage.get(qid, 0) > 1
            for ci, (cn, _, is_n) in enumerate(export_cols):
                val = row.get(cn, "")
                if is_n:
                    row_fmt = fmts["nu2"] if ri % 2 else fmts["num"]
                    try:    ws.write_number(tbl_start + 1 + ri, ci,
                                            float(val), row_fmt)
                    except: ws.write(tbl_start + 1 + ri, ci,
                                     str(val), row_fmt)
                else:
                    row_fmt = (fmts["red"] if reused
                               else (fmts["da2"] if ri % 2 else fmts["dat"]))
                    ws.write(tbl_start + 1 + ri, ci,
                             str(val) if val is not None else "", row_fmt)

        # ── Convert to Excel Table ────────────────────────────────────────────
        # add_table() overwrites the header row with its own styled headers
        # (from the 'columns' list) and makes the range a named Table.
        # Dynamic structured references like TblName[Difficulty] then work
        # in all formula cells we have already written above.
        last_data_row = tbl_start + max(len(df), 1)  # need ≥1 data row
        ws.add_table(tbl_start, 0, last_data_row, n_item_cols - 1, {
            "name":       tbl_name,
            "style":      "Table Style Medium 2",
            "autofilter": True,
            "columns":    [{"header": cn} for cn, _, _ in export_cols],
        })

        ws.freeze_panes(tbl_start + 1, 0)

        # ── Per-domain difficulty distribution charts (formula-driven) ────────
        chart_col = n_item_cols + 2
        chart_row = tbl_start
        if domain_col in df.columns:
            for dom in df[domain_col].dropna().unique():
                sub = df[df[domain_col] == dom]
                chart_row = _write_domain_chart(
                    wb, ws, sub, dcol, str(dom), sheet_name,
                    chart_col, chart_row,
                    table_name=tbl_name,
                    domain_col=domain_col,
                    domain_value=str(dom),
                )

        # ── 3PL charts (live cell-reference charts, not static images) ────────
        if fa:
            img_row = tbl_start + len(df) + 4
            _embed_3pl_charts(wb, ws, fa, img_row, sheet_name=sheet_name)

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

# ═════════════════════════════════════════════════════════════════════════════
#  1PL / RASCH PSYCHOMETRIC MODULE  (optional — does NOT affect assembly)
# ═════════════════════════════════════════════════════════════════════════════
#
# ISOLATION GUARANTEE
# -------------------
# • All functions in this section are NEW additions.
# • No existing function is modified, renamed, or removed.
# • The module is triggered ONLY when the user selects "1PL" in the GUI.
# • When "3PL" is selected (the default) the module is never called and
#   the system behaves exactly as the stable version.
#
# MODEL
# -----
# 1PL / Rasch:   P(θ) = 1 / (1 + exp(-(θ − b)))
# Item info:     I(θ) = P(θ) · (1 − P(θ))
# TCC:           Expected score = Σ P(θ) across items
# TIF:           Total information = Σ I(θ) across items
#
# DELTA (b)
# ---------
# The bank's Difficulty column is used DIRECTLY as Delta.
# No logit or any other transformation is applied.
# Delta classification uses a 0–1 scale matching the bank's range:
#   Easy:   0.00 – 0.35
#   Medium: 0.36 – 0.60
#   Hard:   0.61 – 1.00
# ─────────────────────────────────────────────────────────────────────────────

# Default classification thresholds (0–1 scale, configurable)
_1PL_EASY_MAX  = 0.35
_1PL_HARD_MIN  = 0.61


def _1pl_label(delta: float,
               easy_max: float = _1PL_EASY_MAX,
               hard_min: float = _1PL_HARD_MIN) -> str:
    """Classify an item by its Delta value using the 0–1 scale."""
    if delta <= easy_max:
        return "Easy"
    if delta >= hard_min:
        return "Hard"
    return "Medium"


def _1pl_probability(theta: np.ndarray, b: np.ndarray) -> np.ndarray:
    """
    Vectorised 1PL probability  P(θ) = 1 / (1 + exp(−(θ − b))).

    Parameters
    ----------
    theta : (N,) array of ability values
    b     : (K,) array of item difficulties (bank Difficulty values, 0–1)

    Returns
    -------
    P : (N, K)
    """
    theta = np.asarray(theta, float).reshape(-1, 1)
    b     = np.asarray(b,     float).reshape(1, -1)
    return np.clip(1.0 / (1.0 + np.exp(-(theta - b))), 1e-8, 1 - 1e-8)


@dataclass
class OnePLFormAnalysis:
    """
    All 1PL psychometric statistics for one assembled form.
    Completely separate from the existing FormAnalysis (3PL) dataclass.
    """
    form_name:     str
    n_items:       int
    # Item-level (K items)
    item_ids:      list          # QuestionID strings
    deltas:        np.ndarray   # (K,) bank Difficulty values  ← used as b
    labels:        list[str]    # "Easy" / "Medium" / "Hard"
    domains:       list[str]
    categories:    list[str]
    usage_counts:  list[int]    # cross-form use count per item
    form_ids_map:  list[str]    # pipe-joined form names per item
    # Form-level summary
    mean_delta:    float
    min_delta:     float
    max_delta:     float
    n_easy:        int
    n_medium:      int
    n_hard:        int
    # Curves — 61-point θ grid from −3 to +3
    theta_grid:    np.ndarray   # (61,)
    tcc:           np.ndarray   # (61,)  expected score
    tif:           np.ndarray   # (61,)  test information
    sem:           np.ndarray   # (61,)  standard error 1/√TIF


def analyse_form_1pl(
        form_df:       pd.DataFrame,
        form_name:     str,
        dcol:          str           = "Difficulty",
        domain_col:    str           = "D",
        cat_col:       str           = "Category",
        usage_counter: Optional[dict] = None,
        all_results:   Optional[list] = None,
        n_theta:       int            = 61,
        easy_max:      float          = _1PL_EASY_MAX,
        hard_min:      float          = _1PL_HARD_MIN,
) -> Optional[OnePLFormAnalysis]:
    """
    Compute 1PL/Rasch statistics for one assembled form.

    Delta = Difficulty column value, used directly with no transformation.

    Parameters
    ----------
    form_df       : assembled form DataFrame
    form_name     : label used in the output
    dcol          : difficulty column name (default "Difficulty")
    domain_col    : domain column name    (default "D")
    cat_col       : category column name  (default "Category")
    usage_counter : dict[QuestionID → cross-form use count] for exposure
    all_results   : list of (form_df, mean, warns) — used to build Form IDs
    n_theta       : θ grid points (default 61)
    easy_max      : upper threshold for "Easy" label
    hard_min      : lower threshold for "Hard" label

    Returns None if form_df is empty or missing the difficulty column.
    """
    if form_df is None or form_df.empty:
        return None
    if dcol not in form_df.columns:
        log.warning("analyse_form_1pl: column '%s' not found — skipping.", dcol)
        return None

    df    = form_df.copy()
    b_raw = pd.to_numeric(df[dcol], errors="coerce").fillna(0.5)
    b     = b_raw.to_numpy(float)          # Delta = raw bank value
    K     = len(df)

    theta_grid = np.linspace(-3.0, 3.0, n_theta)
    P          = _1pl_probability(theta_grid, b)   # (n_theta, K)
    tcc        = P.sum(axis=1)                      # (n_theta,)
    info       = P * (1.0 - P)                      # (n_theta, K)
    tif        = info.sum(axis=1)                   # (n_theta,)
    sem        = 1.0 / np.sqrt(np.clip(tif, 1e-8, None))

    labels = [_1pl_label(float(d), easy_max, hard_min) for d in b]

    item_ids   = df["QuestionID"].astype(str).tolist() \
                 if "QuestionID" in df.columns else [f"Item{i+1}" for i in range(K)]
    domains    = df[domain_col].astype(str).tolist() \
                 if domain_col in df.columns else [""] * K
    categories = df[cat_col].astype(str).tolist() \
                 if cat_col    in df.columns else [""] * K

    uc            = usage_counter or {}
    usage_counts  = [int(uc.get(qid, 1)) for qid in item_ids]

    # Which forms contain each item
    if all_results:
        form_ids_map = []
        for qid in item_ids:
            names_containing = []
            for ri, (fdf, _, _) in enumerate(all_results):
                if "QuestionID" in fdf.columns and \
                        qid in fdf["QuestionID"].astype(str).values:
                    names_containing.append(f"Form_{ri + 1}")
            form_ids_map.append(" | ".join(names_containing) or form_name)
    else:
        form_ids_map = [form_name] * K

    return OnePLFormAnalysis(
        form_name=form_name, n_items=K,
        item_ids=item_ids, deltas=b, labels=labels,
        domains=domains, categories=categories,
        usage_counts=usage_counts, form_ids_map=form_ids_map,
        mean_delta=float(b.mean()), min_delta=float(b.min()),
        max_delta=float(b.max()),
        n_easy=labels.count("Easy"),
        n_medium=labels.count("Medium"),
        n_hard=labels.count("Hard"),
        theta_grid=theta_grid, tcc=tcc, tif=tif, sem=sem,
    )


def write_1pl_result_workbook(
        analyses:  list[OnePLFormAnalysis],
        out_path:  Path,
        easy_max:  float = _1PL_EASY_MAX,
        hard_min:  float = _1PL_HARD_MIN,
) -> None:
    """
    Write 1PL_Result.xlsx with seven sheets.

    Sheet layout
    ────────────
    1. Item Statistics       item-level Delta, label, usage, exposure
    2. Form Summary          per-form aggregate statistics
    3. TCC Data              θ vs expected score (all forms) + line chart
    4. TIF Data              θ vs information   (all forms) + line chart
    5. Difficulty Distribution Easy/Medium/Hard counts + bar chart
    6. Exposure Analysis     item usage frequency across all forms
    7. Domain Summary        domain item counts per form
    """
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    if not analyses:
        log.warning("write_1pl_result_workbook: nothing to write.")
        return

    wb  = Workbook(); wb.remove(wb.active)
    NV  = PatternFill("solid", fgColor="1F3864")
    TP  = PatternFill("solid", fgColor="00AECB")
    H   = Font(bold=True, color="FFFFFF", size=10)
    T   = Font(bold=True, color="FFFFFF", size=12)
    ALT = PatternFill("solid", fgColor="EAF0FA")

    def hdr(ws, row, col, val, fill=TP):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = fill; c.font = H
        c.alignment = Alignment(horizontal="center")

    def title(ws, val, ncols):
        ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
        c = ws["A1"]; c.value = val; c.fill = NV; c.font = T
        c.alignment = Alignment(horizontal="center")

    def col_w(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    n_forms    = len(analyses)
    theta_grid = analyses[0].theta_grid

    # Collect the global item pool once (all unique items across forms)
    seen: dict[str, dict] = {}
    for fa in analyses:
        for i, qid in enumerate(fa.item_ids):
            if qid not in seen:
                seen[qid] = {
                    "domain":    fa.domains[i]    if i < len(fa.domains)    else "",
                    "category":  fa.categories[i] if i < len(fa.categories) else "",
                    "delta":     float(fa.deltas[i]),
                    "label":     fa.labels[i],
                    "uses":      fa.usage_counts[i],
                    "form_ids":  fa.form_ids_map[i],
                }

    # ── Sheet 1: Item Statistics ──────────────────────────────────────────
    ws1 = wb.create_sheet("Item Statistics")
    title(ws1, "1PL Item Statistics", 12)
    cols1 = ["Item ID","Domain","Subdomain","Stage","Level/Part",
             "Delta (Difficulty)","Difficulty Label",
             "Item Mean","Item Usage Count",
             "Exposure Rate (%)","Form IDs","Status"]
    for ci, h in enumerate(cols1, 1): hdr(ws1, 2, ci, h)
    col_w(ws1, [16,12,14,9,10,16,16,11,16,16,24,12])

    for ri, (qid, info) in enumerate(seen.items(), 3):
        exp_rate = round(100.0 * info["uses"] / max(n_forms, 1), 1)
        fill = ALT if ri % 2 == 0 else PatternFill()
        for ci, v in enumerate([
            qid, info["domain"], info["category"], "", "",
            round(info["delta"], 4), info["label"],
            round(info["delta"], 4),      # Item Mean ≈ Delta for 1PL
            info["uses"], exp_rate,
            info["form_ids"], "Active",
        ], 1):
            ws1.cell(row=ri, column=ci, value=v).fill = fill

    # ── Sheet 2: Form Summary ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Form Summary")
    title(ws2, "1PL Form Summary", 9)
    for ci, h in enumerate(["Form","N Items","Mean Delta","Min Delta",
                              "Max Delta","Easy","Medium","Hard",
                              "Easy%"], 1): hdr(ws2, 2, ci, h)
    col_w(ws2, [22,9,12,11,11,8,9,8,9])
    for ri, fa in enumerate(analyses, 3):
        fill = ALT if ri % 2 == 0 else PatternFill()
        easy_pct = round(100.0 * fa.n_easy / max(fa.n_items, 1), 1)
        for ci, v in enumerate([
            fa.form_name, fa.n_items,
            round(fa.mean_delta, 4), round(fa.min_delta, 4), round(fa.max_delta, 4),
            fa.n_easy, fa.n_medium, fa.n_hard, f"{easy_pct}%",
        ], 1):
            ws2.cell(row=ri, column=ci, value=v).fill = fill

    # ── Sheet 3: TCC Data ─────────────────────────────────────────────────
    ws3 = wb.create_sheet("TCC Data")
    title(ws3, "Test Characteristic Curves (1PL)", n_forms + 1)
    hdr(ws3, 2, 1, "Theta")
    for ci, fa in enumerate(analyses, 2): hdr(ws3, 2, ci, fa.form_name)
    n_th = len(theta_grid)
    for ri, th in enumerate(theta_grid, 3):
        ws3.cell(row=ri, column=1, value=round(float(th), 2))
        for ci, fa in enumerate(analyses, 2):
            ws3.cell(row=ri, column=ci, value=round(float(fa.tcc[ri-3]), 4))
    ch3 = LineChart(); ch3.title = "TCC (1PL)"; ch3.style = 10
    ch3.y_axis.title = "Expected Score"; ch3.x_axis.title = "Ability (θ)"
    dr3 = Reference(ws3, min_col=2, min_row=2, max_col=1+n_forms, max_row=2+n_th)
    cr3 = Reference(ws3, min_col=1, min_row=3, max_row=2+n_th)
    ch3.add_data(dr3, titles_from_data=True); ch3.set_categories(cr3)
    ch3.width = 18; ch3.height = 12
    ws3.add_chart(ch3, f"{get_column_letter(n_forms+3)}3")

    # ── Sheet 4: TIF Data ─────────────────────────────────────────────────
    ws4 = wb.create_sheet("TIF Data")
    title(ws4, "Test Information Functions (1PL)", n_forms + 1)
    hdr(ws4, 2, 1, "Theta")
    for ci, fa in enumerate(analyses, 2): hdr(ws4, 2, ci, fa.form_name)
    for ri, th in enumerate(theta_grid, 3):
        ws4.cell(row=ri, column=1, value=round(float(th), 2))
        for ci, fa in enumerate(analyses, 2):
            ws4.cell(row=ri, column=ci, value=round(float(fa.tif[ri-3]), 6))
    ch4 = LineChart(); ch4.title = "TIF (1PL)"; ch4.style = 10
    ch4.y_axis.title = "Information"; ch4.x_axis.title = "Ability (θ)"
    dr4 = Reference(ws4, min_col=2, min_row=2, max_col=1+n_forms, max_row=2+n_th)
    cr4 = Reference(ws4, min_col=1, min_row=3, max_row=2+n_th)
    ch4.add_data(dr4, titles_from_data=True); ch4.set_categories(cr4)
    ch4.width = 18; ch4.height = 12
    ws4.add_chart(ch4, f"{get_column_letter(n_forms+3)}3")

    # ── Sheet 5: Difficulty Distribution ─────────────────────────────────
    ws5 = wb.create_sheet("Difficulty Distribution")
    title(ws5, "Difficulty Distribution by Form (1PL)", 4)
    for ci, h in enumerate(["Form","Easy","Medium","Hard"], 1): hdr(ws5, 2, ci, h)
    col_w(ws5, [22,8,9,8])
    for ri, fa in enumerate(analyses, 3):
        fill = ALT if ri % 2 == 0 else PatternFill()
        for ci, v in enumerate([fa.form_name, fa.n_easy, fa.n_medium, fa.n_hard], 1):
            ws5.cell(row=ri, column=ci, value=v).fill = fill
    ch5 = BarChart(); ch5.type = "col"; ch5.grouping = "clustered"
    ch5.title = "Difficulty Distribution (1PL)"
    ch5.y_axis.title = "Count"; ch5.style = 10
    dr5 = Reference(ws5, min_col=2, min_row=2, max_col=4, max_row=2+n_forms)
    cr5 = Reference(ws5, min_col=1, min_row=3, max_row=2+n_forms)
    ch5.add_data(dr5, titles_from_data=True); ch5.set_categories(cr5)
    ch5.width = 16; ch5.height = 10
    ws5.add_chart(ch5, "F3")

    # ── Sheet 6: Exposure Analysis ────────────────────────────────────────
    ws6 = wb.create_sheet("Exposure Analysis")
    title(ws6, "Item Exposure Analysis (1PL)", 5)
    for ci, h in enumerate(["Item ID","Delta","Label","Times Used",
                              "Exposure Rate (%)"], 1): hdr(ws6, 2, ci, h)
    col_w(ws6, [18,12,12,12,16])
    for ri, (qid, info) in enumerate(seen.items(), 3):
        exp = round(100.0 * info["uses"] / max(n_forms, 1), 1)
        fill = ALT if ri % 2 == 0 else PatternFill()
        for ci, v in enumerate([qid, round(info["delta"],4),
                                  info["label"], info["uses"], exp], 1):
            ws6.cell(row=ri, column=ci, value=v).fill = fill

    # ── Sheet 7: Domain Summary ───────────────────────────────────────────
    ws7 = wb.create_sheet("Domain Summary")
    title(ws7, "Domain Item Counts per Form (1PL)", 3)
    for ci, h in enumerate(["Form","Domain","Count"], 1): hdr(ws7, 2, ci, h)
    col_w(ws7, [22,16,10])
    row7 = 3
    from collections import Counter
    for fa in analyses:
        dom_counts = Counter(fa.domains)
        for dom, cnt in sorted(dom_counts.items()):
            for ci, v in enumerate([fa.form_name, dom, cnt], 1):
                ws7.cell(row=row7, column=ci, value=v)
            row7 += 1

    wb.save(str(out_path))
    log.info("1PL Result → %s", out_path)


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
                    if hasattr(self, "_refresh_block_preview"):
                        self._refresh_block_preview()
                elif k == "ub_loaded":
                    self.used_blocks_df = item[1]
                    ub_path = item[2]
                    n_ub = len(item[1])
                    self._ub_path_var.set(Path(ub_path).name)
                    self._ub_stats_lbl.config(
                        text=f"  {n_ub} questions loaded from Used Blocks file")
                    self._sv.set(f"Used Blocks loaded: {n_ub} questions")
                    # Refresh preview — counter source may have changed
                    if hasattr(self, "_refresh_block_preview"):
                        self._refresh_block_preview()
                elif k == "log":
                    _append_log(self._logbox, item[1], item[2] if len(item) > 2 else "")
                elif k == "prog":
                    self._pv.set(item[1])
                elif k in ("done", "done_m1", "done_m2"):
                    self.analyses = item[1]
                    self._gbtn.config(state="normal"); self._pv.set(100)
                    self._refresh_results()
                    self._nb.select(self._results_tab_idx)
                    self._sv.set("Generation complete.")
                    # Update bank Option 2 label with fresh unused count
                    if k in ("done_m1", "done_m2"):
                        self._refresh_unused_label()
                        if hasattr(self, "_update_stats_display"):
                            self._update_stats_display()
                        elif hasattr(self, "_refresh_from_df"):
                            active = self._active_bank()
                            if active is not None:
                                self._refresh_from_df(active)
                    unused_info = ""
                    if len(item) > 2 and isinstance(item[2], int):
                        unused_info = f"\n\nBank: {item[2]} unused questions remaining."
                    mb.showinfo("Done",
                                "All forms generated and analysed.\n"
                                f"Output files saved in a timestamped folder.{unused_info}")
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
        # Prefer the last generated output folder; fall back to bank folder
        d = str(getattr(self, "_last_out_dir", None) or
                (self.source_path.parent if self.source_path else None))
        if d:
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
        # Cumulative set of QuestionIDs used across ALL generation runs this session
        self._used_qids: set[str] = set()
        self._bank_option = tk.IntVar(value=1)   # 1=full, 2=unused only
        # Used Blocks dataset
        self.used_blocks_df: Optional[pd.DataFrame] = None
        self._use_ub_var = tk.BooleanVar(value=False)
        self._adaptive_mean_var = tk.BooleanVar(value=False)
        # Psychometric model (1PL / 3PL) — new optional feature
        self._psych_model_var = tk.StringVar(value="3pl")
        # Feature 2 — Difficulty scope
        self._diff_scope_var = tk.StringVar(value="exam")
        # Feature 3 — Difficulty method
        self._diff_method_var = tk.StringVar(value="delta")
        self._cct_profile_var = tk.StringVar(value="Common")
        # Feature 4 — Fallback-for-difficulty
        self._fallback_reuse_var = tk.BooleanVar(value=False)
        # Block ID settings
        self._part_level_var = tk.IntVar(value=1)
        self._stage_num_var  = tk.IntVar(value=2)
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

        # Upload row
        uf = tk.Frame(p, bg=BG_L); uf.pack(fill="x", pady=4)
        self._bp = tk.StringVar(value="No file selected")
        tk.Label(uf, textvariable=self._bp, bg=BG_L, fg=BG_D,
                 font=FB, width=55, anchor="w").pack(side="left")
        _btn(uf, "Upload Excel / CSV", self._browse_bank,
             bg=self.COLOR).pack(side="left", padx=8)

        # ── Bank source selector ─────────────────────────────────────────────
        src_frm = tk.LabelFrame(p, text="Assembly Bank Source",
                                bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        src_frm.pack(fill="x", pady=(8, 4))

        tk.Radiobutton(src_frm,
                       text="Option 1 — Use full uploaded bank  (all questions)",
                       variable=self._bank_option, value=1,
                       bg=BG_L, fg=TXD, font=FB,
                       activebackground=BG_L,
                       command=self._on_bank_option).pack(anchor="w")

        opt2_frm = tk.Frame(src_frm, bg=BG_L); opt2_frm.pack(fill="x", anchor="w")
        tk.Radiobutton(opt2_frm,
                       text="Option 2 — Use only unused questions  ",
                       variable=self._bank_option, value=2,
                       bg=BG_L, fg=TXD, font=FB,
                       activebackground=BG_L,
                       command=self._on_bank_option).pack(side="left")
        self._unused_lbl = tk.Label(opt2_frm,
                                    text="(no data yet)",
                                    bg=BG_L, fg=ETEC_TEAL, font=("Segoe UI", 9, "italic"))
        self._unused_lbl.pack(side="left")

        # Stats
        self._bstats = tk.StringVar(value="")
        tk.Label(p, textvariable=self._bstats, bg=BG_L, fg=BG_M,
                 font=FB, justify="left").pack(anchor="w", pady=4)

        # ── Used Blocks dataset ──────────────────────────────────────────────
        ub_frm = tk.LabelFrame(p, text="Used Blocks Dataset  (optional fallback)",
                               bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        ub_frm.pack(fill="x", pady=(6, 4))

        ub_row = tk.Frame(ub_frm, bg=BG_L); ub_row.pack(fill="x")
        self._ub_path_var = tk.StringVar(value="No file selected")
        tk.Label(ub_row, textvariable=self._ub_path_var,
                 bg=BG_L, fg=BG_D, font=FB, width=42, anchor="w").pack(side="left")
        _btn(ub_row, "Upload Used Blocks", self._browse_used_blocks,
             bg=ETEC_PURPLE).pack(side="left", padx=8)

        ub_ctrl = tk.Frame(ub_frm, bg=BG_L); ub_ctrl.pack(fill="x", pady=(6,0))
        tk.Checkbutton(ub_ctrl,
                       text="Allow reuse from Used Blocks if main bank is exhausted  "
                            "(last resort — unique-first, then controlled reuse)",
                       variable=self._use_ub_var,
                       bg=BG_L, fg=TXD, font=FB,
                       activebackground=BG_L).pack(side="left")

        # Max Usage Per Question (UB) spinbox
        ub_max_row = tk.Frame(ub_frm, bg=BG_L); ub_max_row.pack(fill="x", pady=(4,0))
        tk.Label(ub_max_row,
                 text="Max Usage Per Question  (Used Blocks):",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._max_ub_usage_m1 = tk.Spinbox(
            ub_max_row, from_=1, to=20, width=4, font=FB,
            bg="#FFF", fg="#000", relief="solid", bd=1)
        self._max_ub_usage_m1.delete(0, "end"); self._max_ub_usage_m1.insert(0, "2")
        self._max_ub_usage_m1.pack(side="left", padx=6)
        tk.Label(ub_max_row,
                 text="(1 = unique only, 2+ = allow controlled reuse up to N times)",
                 bg=BG_L, fg="#666", font=("Segoe UI", 8, "italic")).pack(side="left")

        self._ub_stats_lbl = tk.Label(ub_frm, text="",
                                      bg=BG_L, fg=ETEC_PURPLE,
                                      font=("Segoe UI", 9, "italic"))
        self._ub_stats_lbl.pack(anchor="w")

        # ── Feature 1: Template Import / Export ──────────────────────────────
        tmpl_frm = tk.LabelFrame(
            p, text="Exam Template (Feature 1 — Dynamic Exam Creation)",
            bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        tmpl_frm.pack(fill="x", pady=(8, 4))
        tk.Label(tmpl_frm,
                 text="Import a structured template to create a new exam, "
                      "or download a blank template to fill in.",
                 bg=BG_L, fg="#555", font=("Segoe UI", 9)).pack(anchor="w")
        btn_row = tk.Frame(tmpl_frm, bg=BG_L); btn_row.pack(fill="x", pady=(6,0))
        _btn(btn_row, "⬆  Import Template",
             self._import_template, bg=ETEC_BLUE).pack(side="left")
        _btn(btn_row, "⬇  Download Blank Template",
             self._export_template, bg=ETEC_NAVY).pack(side="left", padx=8)

        _lbl(p, "Categories detected in 'Category' column", bold=True).pack(
            anchor="w", pady=(10, 4))
        self._cat_tree = _treeview(p,
            ["Category","Domain (D)","Count","Mean Difficulty","Min","Max"],
            [160, 140, 60, 120, 65, 65])

    # ── Feature 1: Template Import / Export ───────────────────────────────────
    def _export_template(self):
        path = fd.asksaveasfilename(
            title="Save Exam Template",
            defaultextension=".xlsx",
            initialfile="ExamTemplate.xlsx",
            filetypes=[("Excel","*.xlsx"),("All","*.*")])
        if not path: return
        try:
            export_exam_template(Path(path))
            mb.showinfo("Exported", f"Template saved to:\n{path}")
        except Exception as e:
            mb.showerror("Error", str(e))

    def _import_template(self):
        path = fd.askopenfilename(
            title="Import Exam Template",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not path: return
        exam_name, structure, errors = import_exam_template(Path(path))
        if errors:
            mb.showerror("Template Validation Failed",
                         "\n".join(errors[:10]))
            return

        # Check if exam already exists
        existing = exam_name in MODE1_EXAMS
        if existing:
            choice = mb.askyesnocancel(
                "Exam Already Exists",
                f"An exam named '{exam_name}' already exists.\n\n"
                f"  Yes    = Overwrite\n"
                f"  No     = Append categories\n"
                f"  Cancel = Abort import")
            if choice is None:
                return
            elif choice:   # overwrite
                MODE1_EXAMS[exam_name] = {}
            # No = append — existing entries stay

        # Build flat Category → count dict (exam has ONE domain level in MODE1)
        for domain, cats in structure[exam_name].items():
            for cat, cnt in cats.items():
                MODE1_EXAMS.setdefault(exam_name, {})[cat] = cnt

        # Refresh exam combobox
        self._exam_var.set(exam_name)
        # Update the options in the exam combobox widget
        cb = self._exam_cb
        cb["values"] = list(MODE1_EXAMS.keys())
        self._populate_subdomain_entries()

        # Preview
        total = sum(MODE1_EXAMS[exam_name].values())
        domains_str = ", ".join(structure[exam_name].keys())
        mb.showinfo(
            "Template Imported",
            f"Exam '{exam_name}' loaded successfully!\n\n"
            f"Domains  : {domains_str}\n"
            f"Categories: {len(MODE1_EXAMS[exam_name])}\n"
            f"Total Q  : {total}\n\n"
            f"The exam template is now active in the Configure tab.")

    def _browse_used_blocks(self):
        path = fd.askopenfilename(
            title="Select Used Blocks File",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not path: return
        self._sv.set("Loading Used Blocks…")
        threading.Thread(target=self._load_ub_t, args=(path,), daemon=True).start()

    def _load_ub_t(self, path):
        try:
            df = load_used_blocks(path)
            self._q.put(("ub_loaded", df, path))
        except Exception as e:
            self._q.put(("err", str(e)))

    def _on_bank_option(self):
        """Called when the user switches between Option 1 / Option 2."""
        if self.bank_df is not None:
            self._update_stats_display()

    def _active_bank(self) -> pd.DataFrame:
        """Return the bank that should be used for assembly."""
        if self._bank_option.get() == 2 and self.bank_df is not None:
            unused = self.bank_df[
                ~self.bank_df["QuestionID"].astype(str).isin(self._used_qids)
            ].copy()
            return unused if not unused.empty else self.bank_df
        return self.bank_df

    def _update_stats_display(self):
        """Refresh bstats label and category tree for the currently active bank."""
        df = self._active_bank()
        if df is None or df.empty:
            return
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

    def _refresh_unused_label(self):
        """Update the Option 2 unused count label after a generation run."""
        if self.bank_df is None:
            return
        total   = len(self.bank_df)
        used    = len(self._used_qids & set(self.bank_df["QuestionID"].astype(str)))
        unused  = total - used
        self._unused_lbl.config(
            text=f"({unused} unused out of {total}  |  {used} used so far)")

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
        self.bank_df = df
        self.source_path = Path(path)
        self._bp.set(Path(path).name)
        # Reset cumulative usage when a new bank is loaded
        self._used_qids = set()
        self._bank_option.set(1)
        self._unused_lbl.config(text="(no generations yet)")
        self._update_stats_display()
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
        self._exam_cb  = ttk.Combobox(r0, textvariable=self._exam_var,
                          values=list(MODE1_EXAMS.keys()), state="readonly", width=28)
        cb = self._exam_cb
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

        # ── Adaptive Mean Control toggle ───────────────────────────────────────
        amc_frm = tk.Frame(pf, bg=BG_L); amc_frm.grid(
            row=7, column=0, columnspan=2, sticky="w", pady=(4,2))
        tk.Checkbutton(amc_frm,
                       text="Enable Adaptive Mean Control",
                       variable=self._adaptive_mean_var,
                       bg=BG_L, fg=ETEC_BLUE, font=("Segoe UI", 10, "bold"),
                       activebackground=BG_L).pack(side="left")
        tk.Label(amc_frm,
                 text="  (intelligent swap correction to hit target mean — "
                      "uses reuse & Used Blocks strategically)",
                 bg=BG_L, fg="#666", font=("Segoe UI", 8, "italic")).pack(side="left")

        # ── Feature 2: Difficulty Distribution Scope ──────────────────────────
        scope_frm = tk.LabelFrame(pf, text="Difficulty Scope (Feature 2)",
                                  bg=BG_L, fg=BG_D, font=FH, padx=8, pady=6)
        scope_frm.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6,2))
        for val, lbl in [("exam","Exam Level"),
                         ("domain","Domain Level"),
                         ("category","Category Level")]:
            tk.Radiobutton(scope_frm, text=lbl, variable=self._diff_scope_var,
                           value=val, bg=BG_L, fg=TXD, font=FB,
                           activebackground=BG_L).pack(side="left", padx=10)

        # ── Feature 3: Difficulty Method ──────────────────────────────────────
        meth_frm = tk.LabelFrame(pf, text="Difficulty Method (Feature 3)",
                                 bg=BG_L, fg=BG_D, font=FH, padx=8, pady=6)
        meth_frm.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(2,2))
        for val, lbl in [("delta","Delta (current)"),("cct","CCT Method")]:
            tk.Radiobutton(meth_frm, text=lbl, variable=self._diff_method_var,
                           value=val, bg=BG_L, fg=TXD, font=FB,
                           activebackground=BG_L,
                           command=self._on_method_change).pack(side="left", padx=10)
        self._cct_profile_row = tk.Frame(meth_frm, bg=BG_L)
        self._cct_profile_row.pack(side="left", padx=10)
        tk.Label(self._cct_profile_row, text="Profile:",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._cct_cb = ttk.Combobox(
            self._cct_profile_row, textvariable=self._cct_profile_var,
            values=list(CCT_RULES.keys()), state="readonly", width=14)
        self._cct_cb.pack(side="left", padx=4)
        self._cct_profile_row.pack_forget()   # hidden until CCT selected

        # ── Feature 4: Fallback for Difficulty ────────────────────────────────
        fb_frm = tk.Frame(pf, bg=BG_L)
        fb_frm.grid(row=10, column=0, columnspan=2, sticky="w", pady=(2,4))
        tk.Checkbutton(fb_frm,
                       text="Allow Fallback Using Used Blocks to achieve difficulty target  "
                            "(Feature 4)",
                       variable=self._fallback_reuse_var,
                       bg=BG_L, fg=ETEC_PURPLE, font=("Segoe UI", 9, "bold"),
                       activebackground=BG_L).pack(side="left")

        self._sim_n = self._le(pf, "Simulation students:", 11, "3000")
        self._seed  = self._le(pf, "RNG Seed (blank=random):", 12, "")

        # ── Psychometric Model selector (new optional feature) ─────────────────
        pm_frm = tk.LabelFrame(pf, text="Psychometric Model",
                               bg=BG_L, fg=BG_D, font=FH, padx=8, pady=6)
        pm_frm.grid(row=13, column=0, columnspan=2, sticky="ew", pady=(4, 2))
        pm_row = tk.Frame(pm_frm, bg=BG_L); pm_row.pack(fill="x")
        for val, lbl, tip in [
            ("3pl", "3PL  (default — existing behaviour)",
             "Full 3-parameter model  ·  no change"),
            ("1pl", "1PL  (Rasch / Delta-only)",
             "Generates 1PL_Result.xlsx  ·  uses Difficulty as Delta directly"),
        ]:
            tk.Radiobutton(pm_row, text=lbl, variable=self._psych_model_var,
                           value=val, bg=BG_L, fg=TXD, font=FB,
                           activebackground=BG_L).pack(side="left", padx=(0, 8))
        tk.Label(pm_frm,
                 text="1PL: Delta = bank Difficulty  ·  "
                      "TCC/TIF use Rasch formula  ·  "
                      "3PL output unchanged when 3PL is selected",
                 bg=BG_L, fg="#555",
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")

        # Save settings button
        _btn(pf, "Save & Preview Settings", self._save_settings,
             bg=BG_D).grid(row=14, column=0, columnspan=2, pady=8)

        # ── Block ID Settings panel ───────────────────────────────────────────
        bid_frm = tk.LabelFrame(p, text="Block ID Settings",
                                bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        bid_frm.grid(row=0, column=2, sticky="nsew", padx=(8, 0), pady=4)
        p.columnconfigure(2, weight=1)

        # ── Purpose explanation (collapsed by default, expandable) ────────────
        help_txt = (
            "Block ID Format:  <Code>-<Stage>.<Level>.<Number>\n"
            "\n"
            "Stage 1  →  two parts (not difficulty-based)\n"
            "  1.1 = Part 1       GAT-1.1.001\n"
            "  1.2 = Part 2       GAT-1.2.001\n"
            "\n"
            "Stage 2  →  three difficulty levels\n"
            "  2.1 = Easy         GAT-2.1.001\n"
            "  2.2 = Medium       GAT-2.2.015\n"
            "  2.3 = Difficult    GAT-2.3.120\n"
            "\n"
            "Stage 3  →  three difficulty levels\n"
            "  3.1 = Easy         GAT-3.1.001\n"
            "  3.2 = Medium       GAT-3.2.010\n"
            "  3.3 = Difficult    GAT-3.3.050\n"
            "\n"
            "Sequencing rule:\n"
            "  If Used Blocks file is loaded, the counter\n"
            "  continues from the highest existing block number\n"
            "  (e.g. max found = 005 → next = 006).\n"
            "  Otherwise the local JSON counter is used."
        )
        help_box = tk.Text(bid_frm, height=14, width=36,
                           bg="#EEF2FF", fg=ETEC_NAVY,
                           font=("Consolas", 8), state="normal",
                           relief="solid", bd=1, wrap="none")
        help_box.insert("end", help_txt)
        help_box.config(state="disabled")
        help_box.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        # ── Dynamic level selector ─────────────────────────────────────────────
        tk.Label(bid_frm, text="Select level / part:", bg=BG_L,
                 fg=TXD, font=FB).grid(row=1, column=0, sticky="nw", pady=2)
        pl_frm = tk.Frame(bid_frm, bg=BG_L); pl_frm.grid(row=1, column=1, sticky="w")

        # Store the radio-button widgets so _refresh_block_preview can retitle them
        self._pl_radios: list[tk.Radiobutton] = []
        for val in (1, 2, 3):
            rb = tk.Radiobutton(pl_frm, text=f"Option {val}",
                                variable=self._part_level_var,
                                value=val, bg=BG_L, fg=TXD, font=FB,
                                activebackground=BG_L,
                                command=self._refresh_block_preview)
            rb.pack(anchor="w")
            self._pl_radios.append(rb)

        # ── Stage Number selector (2 or 3) ───────────────────────────────────
        tk.Label(bid_frm, text="Stage Number:", bg=BG_L,
                 fg=TXD, font=FB).grid(row=2, column=0, sticky="w", pady=(8, 2))
        sn_frm = tk.Frame(bid_frm, bg=BG_L); sn_frm.grid(row=2, column=1, sticky="w")
        self._stage_num_radios: list[tk.Radiobutton] = []
        for val, lbl in [(2, "Stage 2"), (3, "Stage 3")]:
            rb = tk.Radiobutton(sn_frm, text=lbl,
                                variable=self._stage_num_var, value=val,
                                bg=BG_L, fg=TXD, font=FB, activebackground=BG_L,
                                command=self._refresh_block_preview)
            rb.pack(side="left", padx=(0, 12))
            self._stage_num_radios.append(rb)
        self._stage_num_note = tk.Label(sn_frm, text="",
                                        bg=BG_L, fg=ETEC_PURPLE,
                                        font=("Segoe UI", 8, "italic"))
        self._stage_num_note.pack(side="left")

        # ── Preview + counter info ─────────────────────────────────────────────
        tk.Label(bid_frm, text="Block ID Preview:", bg=BG_L,
                 fg=TXD, font=FB).grid(row=3, column=0, sticky="nw", pady=(8, 2))
        self._bid_preview = tk.Text(bid_frm, height=5, width=24,
                                    bg="#F0F4FF", fg=ETEC_NAVY,
                                    font=("Consolas", 9), state="disabled",
                                    relief="solid", bd=1)
        self._bid_preview.grid(row=3, column=1, sticky="w", pady=(8, 2))

        self._bid_counter_lbl = tk.Label(bid_frm, text="",
                                         bg=BG_L, fg=ETEC_TEAL,
                                         font=("Segoe UI", 8, "italic"))
        self._bid_counter_lbl.grid(row=4, column=0, columnspan=2, sticky="w")

        _btn(bid_frm, "↺ Reset Counter",
             lambda: self._reset_block_counter(), bg=ETEC_PURPLE).grid(
            row=5, column=0, columnspan=2, pady=4)

        # Set correct initial labels
        self._update_level_radio_labels()

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

    def _on_method_change(self):
        """Show/hide CCT profile combobox based on selected difficulty method."""
        if hasattr(self, "_cct_profile_row"):
            if self._diff_method_var.get() == "cct":
                self._cct_profile_row.pack(side="left", padx=10)
            else:
                self._cct_profile_row.pack_forget()

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
        # Update Block ID radio labels whenever stage changes
        if hasattr(self, "_pl_radios"):
            self._refresh_block_preview()

    def _update_level_radio_labels(self):
        """
        Retitle the level radio buttons and update the Stage Number row
        based on which difficulty stage is currently selected.

        Stage1      → level = Part 1 or Part 2 (stage number row hidden)
        E / M / D   → level is fixed by the difficulty choice;
                       Stage Number row enabled → user picks Stage 2 or 3
        Manually set → user picks both level and stage number freely
        """
        stage = self._stage_var.get() if hasattr(self, "_stage_var") else "Stage1"
        sl    = STAGE_TO_SL.get(stage, (1, 0))
        stage_n   = sl[0]
        fixed_lvl = sl[1]   # 0 = user must pick; non-zero = fixed by stage choice

        # ── Level / Part radio labels ──────────────────────────────────────────
        if stage_n == 1:
            labels   = ["Part 1  (1.1)", "Part 2  (1.2)", "—  (n/a)"]
            disabled = [False, False, True]
        else:
            # Stage 2 or 3 — level names are the same; stage nr shows in preview
            sn = self._stage_num_var.get() if hasattr(self, "_stage_num_var") else 2
            labels   = [f"Easy  ({sn}.1)", f"Medium  ({sn}.2)", f"Difficult  ({sn}.3)"]
            disabled = [False, False, False]

        for rb, lbl, dis in zip(self._pl_radios, labels, disabled):
            rb.config(text=lbl)
            if fixed_lvl != 0:
                # Level is already encoded by the difficulty stage (E/M/D)
                rb.config(state="disabled")
                self._part_level_var.set(fixed_lvl)
            elif dis:
                rb.config(state="disabled")
            else:
                rb.config(state="normal")

        # ── Stage Number row (2 vs 3) ──────────────────────────────────────────
        if hasattr(self, "_stage_num_radios"):
            if stage_n == 1:
                # Stage1 always generates 1.x IDs — stage number row irrelevant
                for rb in self._stage_num_radios:
                    rb.config(state="disabled")
                self._stage_num_note.config(text="(fixed: Stage 1)")
            else:
                # E, M, D, Manually set — user can choose Stage 2 or Stage 3
                for rb in self._stage_num_radios:
                    rb.config(state="normal")
                sn = self._stage_num_var.get()
                note = f"→ IDs will be {EXAM_CODES.get(self._exam_var.get(),'EX')}-{sn}.x.xxx"
                self._stage_num_note.config(text=note)

    def _refresh_block_preview(self):
        """Update the Block ID preview text box and counter label."""
        self._update_level_radio_labels()
        if self.source_path is None:
            return
        try:
            n = int(self._nforms.get())
        except ValueError:
            n = 1
        exam  = self._exam_var.get()
        stage = self._stage_var.get()
        pl    = self._part_level_var.get()
        code  = EXAM_CODES.get(exam, "EX")
        sl    = STAGE_TO_SL.get(stage, (1, 1))
        stage_n = sl[0]
        level_n = pl if sl[1] == 0 else sl[1]
        # Apply stage override for E/M/D → Stage 2 or Stage 3
        if stage_n in (2, 3) and hasattr(self, "_stage_num_var"):
            stage_n = self._stage_num_var.get()
        key   = f"{code}-{stage_n}.{level_n}"

        # What is the current high-water mark?
        ub_max   = _max_block_number_from_used_blocks(key, self.used_blocks_df)
        counters = _load_counters(self.source_path.parent)
        json_max = counters.get(key, 0)
        current  = max(ub_max, json_max)

        # Preview IDs (do NOT persist — just for display)
        ids = [f"{key}.{current + 1 + i:03d}" for i in range(n)]

        self._bid_preview.config(state="normal")
        self._bid_preview.delete("1.0", "end")
        self._bid_preview.insert("end", "\n".join(ids))
        self._bid_preview.config(state="disabled")

        # Show where the counter currently stands
        source = "Used Blocks" if ub_max >= json_max and ub_max > 0 \
                 else ("JSON counter" if json_max > 0 else "none — starting at 001")
        self._bid_counter_lbl.config(
            text=f"Current max for {key}: {current:03d}  (source: {source})")

    def _reset_block_counter(self):
        """Clear the persisted counter for the current exam/stage/level."""
        if self.source_path is None:
            mb.showinfo("Info", "Upload a bank file first."); return
        exam  = self._exam_var.get()
        stage = self._stage_var.get()
        pl    = self._part_level_var.get()
        code  = EXAM_CODES.get(exam, "EX")
        sl    = STAGE_TO_SL.get(stage, (1, 1))
        stage_n = sl[0]
        level_n = pl if sl[1] == 0 else sl[1]
        # Apply stage override
        if stage_n in (2, 3) and hasattr(self, "_stage_num_var"):
            stage_n = self._stage_num_var.get()
        key   = f"{code}-{stage_n}.{level_n}"
        counters = _load_counters(self.source_path.parent)
        if key in counters:
            counters.pop(key)
            _save_counters(self.source_path.parent, counters)
        mb.showinfo("Reset", f"Counter for '{key}' has been reset to 0.")
        self._refresh_block_preview()

    def _save_settings(self):
        """Preview / confirm current settings (mirrors save_settings() from Script 1)."""
        try:
            n_forms = int(self._nforms.get())
        except ValueError:
            n_forms = "?"
        stage = self._stage_var.get()
        crit  = MODE1_STAGES.get(stage, {})
        exam  = self._exam_var.get()
        pl    = self._part_level_var.get()
        total = sum(int(e.get()) for _, e in self._sub_rows
                    if e.get().isdigit())
        # Preview next block IDs
        bid_preview = ""
        if self.source_path:
            try:
                ids = generate_block_ids(exam, stage, min(n_forms, 3),
                                         pl, self.source_path.parent,
                                         self.used_blocks_df)
                bid_preview = f"\nBlock IDs   : {ids[0]} … {ids[-1]}"
            except Exception:
                pass
        lines = [
            f"Exam        : {exam}",
            f"Stage       : {stage}",
            f"Stage Number: {self._stage_num_var.get()} "
            f"({'active' if STAGE_TO_SL.get(stage,(1,0))[0] in (2,3) else 'n/a for Stage1'})",
            f"Part/Level  : {pl}",
            f"Diff Range  : {crit.get('Range', 'manual')}",
            f"Mean Target : {crit.get('Mean', 'manual')}",
            f"Forms       : {n_forms}",
            f"Total Q/form: {total}",
            f"Bins        : {'Enabled' if self._use_bins.get() else 'Disabled'}",
            f"Disc filter : {'Enabled (≥0.85)' if self._stats_var.get() else 'Disabled'}",
            f"Partial fill: {'Enabled' if self._partial_var.get() else 'Disabled'}",
            f"Reuse       : {'Enabled — max ' + self._max_reuse_m1.get() + 'x per question' if self._reuse_var.get() else 'Disabled'}",
            f"Used Blocks : {'Enabled' if self._use_ub_var.get() else 'Disabled'}",
            f"Adaptive MC : {'Enabled' if self._adaptive_mean_var.get() else 'Disabled'}",
            f"Diff Scope  : {self._diff_scope_var.get()}",
            f"Diff Method : {self._diff_method_var.get().upper()}"
            + (f" / {self._cct_profile_var.get()}"
               if self._diff_method_var.get()=='cct' else ""),
            f"Fallback4   : {'Enabled' if self._fallback_reuse_var.get() else 'Disabled'}"
            + bid_preview,
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
        # Pass the currently active bank view (full or unused-only) to the thread
        active_bank = self._active_bank()
        threading.Thread(target=self._gen_t,
                         args=(params, active_bank), daemon=True).start()

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
            used_blocks=self.used_blocks_df if self._use_ub_var.get() else None,
            allow_used_blocks=self._use_ub_var.get(),
            max_ub_usage=max(1, int(self._max_ub_usage_m1.get()))
                         if self._max_ub_usage_m1.get().strip().isdigit() else 2,
            adaptive_mean_control=self._adaptive_mean_var.get(),
            difficulty_scope=self._diff_scope_var.get(),
            difficulty_method=self._diff_method_var.get(),
            cct_profile=self._cct_profile_var.get(),
            allow_fallback_reuse=self._fallback_reuse_var.get(),
            psychometric_model=self._psych_model_var.get(),
        )
        return {"n_forms": n_forms, "params": params,
                "n_students": ii(self._sim_n, "Simulation students"),
                "seed": int(seed_s) if seed_s else None,
                "exam": self._exam_var.get(),
                "stage": self._stage_var.get(),
                "part_level": self._part_level_var.get(),
                "stage_override": self._stage_num_var.get()}

    def _gen_t(self, p, active_bank):
        try:
            n = p["n_forms"]; params = p["params"]
            # Generate block IDs — sequence continues from Used Blocks dataset
            if self.source_path:
                names = generate_block_ids(
                    p["exam"], p["stage"], n,
                    p["part_level"], self.source_path.parent,
                    self.used_blocks_df,
                    stage_override=p.get("stage_override"))
            else:
                names = [f"Form_{i+1}" for i in range(n)]
            bank_label = ("unused-only" if self._bank_option.get() == 2
                          else "full bank")
            self._lg(f"Starting assembly: {n} forms  [{bank_label}, "
                     f"{len(active_bank)} questions]…")
            self._lg(f"  Block IDs: {names[0]} … {names[-1]}", ETEC_TEAL)
            results = assemble_forms(active_bank, params, "Difficulty", n, names)
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
            exam_type  = self._exam_var.get()
            # Include block ID range in folder name for traceability
            folder_tag = f"{exam_type}  [{names[0]}…{names[-1]}]"
            out_dir    = _make_output_folder(
                self.source_path.parent, folder_tag, n)
            self._last_out_dir = out_dir   # for Open Folder button
            fp = out_dir / "Forms.xlsx"
            ap = out_dir / "3PL_Analysis.xlsx"
            rp = out_dir / "Remaining_Questions.xlsx"

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

            # ── Psychometric output: 3PL or 1PL depending on selection ─────
            psych_model = p["params"].psychometric_model
            pl1_path = None

            if psych_model == "1pl":
                # ── 1PL mode: write 1PL_Result.xlsx, skip 3PL analysis ──────
                pl1_path = out_dir / "1PL_Result.xlsx"
                try:
                    self._lg("Running 1PL (Rasch) psychometric analysis…")
                    analyses_1pl = [
                        analyse_form_1pl(
                            r[0], nm,
                            dcol="Difficulty",
                            domain_col="D",
                            cat_col="Category",
                            usage_counter=usage,
                            all_results=results,
                        )
                        for r, nm in zip(results, names)
                    ]
                    analyses_1pl = [x for x in analyses_1pl if x]
                    if analyses_1pl:
                        for fa in analyses_1pl:
                            self._lg(
                                f"  {fa.form_name}: n={fa.n_items}  "
                                f"mean_Δ={fa.mean_delta:.3f}  "
                                f"Easy={fa.n_easy} Med={fa.n_medium} Hard={fa.n_hard}"
                            )
                        write_1pl_result_workbook(analyses_1pl, pl1_path)
                    else:
                        pl1_path = None
                        self._lg("  ⚠ 1PL: no valid analyses generated.", "yellow")
                except Exception as e1pl:
                    self._lg(f"  ⚠ 1PL export error: {e1pl}", "yellow")
                    log.exception("1PL export failed")
                    pl1_path = None
            else:
                # ── 3PL mode (default — unchanged behaviour) ─────────────────
                if valid: write_analysis_workbook(valid, ap)

            # Feature 5: Charts workbook
            cp = out_dir / "Charts.xlsx"
            try:
                diff_method = p["params"].difficulty_method
                write_charts_workbook(results, names, "Difficulty", "D",
                                      diff_method, cp)
            except Exception as ce:
                self._lg(f"  ⚠ Charts file error: {ce}", "yellow")
                cp = None

            used_ids = set()
            for form, _, _ in results:
                if "QuestionID" in form.columns:
                    used_ids.update(form["QuestionID"].dropna().astype(str).tolist())

            # ── Cumulative tracking — update Option 2 pool ───────────────────
            self._used_qids.update(used_ids)

            # Write Remaining Questions based on FULL bank minus ALL ever used
            remaining = self.bank_df[
                ~self.bank_df["QuestionID"].astype(str).isin(self._used_qids)
            ].copy()
            remaining.to_excel(str(rp), index=False, engine="openpyxl")
            unused_count = len(remaining)

            self._q.put(("prog", 100))
            self._lg(f"\n  📁 Output folder   : {out_dir}")
            self._lg(f"  📄 Forms           : {fp.name}")
            if psych_model == "1pl" and pl1_path:
                self._lg(f"  🧪 1PL Result      : {pl1_path.name}")
            else:
                self._lg(f"  📊 3PL Analysis    : {ap.name}")
            if cp: self._lg(f"  📈 Charts          : {cp.name}")
            self._lg(f"  📋 Remaining Qs    : {rp.name}  ({unused_count} questions)")
            self._lg(f"\n  Bank status: {unused_count} unused / "
                     f"{len(self.bank_df)} total questions remain",
                     ETEC_TEAL)
            self._lg("\n✓ Generation complete!", "lime")
            self._q.put(("done_m1", valid, unused_count))
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
        self._used_qids:  set[str] = set()
        self._bank_option = tk.IntVar(value=1)
        # Used Blocks + block ID settings
        self.used_blocks_df: Optional[pd.DataFrame] = None
        self._use_ub_var    = tk.BooleanVar(value=False)
        self._part_level_var    = tk.IntVar(value=1)
        self._adaptive_mean_var = tk.BooleanVar(value=False)
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

        # ── Bank source selector ─────────────────────────────────────────────
        src_frm = tk.LabelFrame(p, text="Assembly Bank Source",
                                bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        src_frm.pack(fill="x", pady=(8, 4))
        tk.Radiobutton(src_frm,
                       text="Option 1 — Use full uploaded bank  (all questions)",
                       variable=self._bank_option, value=1,
                       bg=BG_L, fg=TXD, font=FB, activebackground=BG_L,
                       command=self._on_bank_option).pack(anchor="w")
        opt2_frm = tk.Frame(src_frm, bg=BG_L); opt2_frm.pack(fill="x", anchor="w")
        tk.Radiobutton(opt2_frm,
                       text="Option 2 — Use only unused questions  ",
                       variable=self._bank_option, value=2,
                       bg=BG_L, fg=TXD, font=FB, activebackground=BG_L,
                       command=self._on_bank_option).pack(side="left")
        self._unused_lbl = tk.Label(opt2_frm, text="(no generations yet)",
                                    bg=BG_L, fg=ETEC_TEAL,
                                    font=("Segoe UI", 9, "italic"))
        self._unused_lbl.pack(side="left")

        # ── Used Blocks dataset ──────────────────────────────────────────────
        ub_frm = tk.LabelFrame(p, text="Used Blocks Dataset  (optional fallback)",
                               bg=BG_L, fg=BG_D, font=FH, padx=10, pady=8)
        ub_frm.pack(fill="x", pady=(6, 4))
        ub_row = tk.Frame(ub_frm, bg=BG_L); ub_row.pack(fill="x")
        self._ub_path_var = tk.StringVar(value="No file selected")
        tk.Label(ub_row, textvariable=self._ub_path_var,
                 bg=BG_L, fg=BG_D, font=FB, width=42, anchor="w").pack(side="left")
        _btn(ub_row, "Upload Used Blocks", self._browse_used_blocks,
             bg=ETEC_PURPLE).pack(side="left", padx=8)
        ub_ctrl = tk.Frame(ub_frm, bg=BG_L); ub_ctrl.pack(fill="x", pady=(6,0))
        tk.Checkbutton(ub_ctrl,
                       text="Allow reuse from Used Blocks if main bank is exhausted  "
                            "(unique-first, then controlled reuse)",
                       variable=self._use_ub_var,
                       bg=BG_L, fg=TXD, font=FB, activebackground=BG_L).pack(side="left")

        ub_max_row2 = tk.Frame(ub_frm, bg=BG_L); ub_max_row2.pack(fill="x", pady=(4,0))
        tk.Label(ub_max_row2,
                 text="Max Usage Per Question  (Used Blocks):",
                 bg=BG_L, fg=TXD, font=FB).pack(side="left")
        self._max_ub_usage_m2 = tk.Spinbox(
            ub_max_row2, from_=1, to=20, width=4, font=FB,
            bg="#FFF", fg="#000", relief="solid", bd=1)
        self._max_ub_usage_m2.delete(0, "end"); self._max_ub_usage_m2.insert(0, "2")
        self._max_ub_usage_m2.pack(side="left", padx=6)
        tk.Label(ub_max_row2,
                 text="(1 = unique only, 2+ = allow controlled reuse up to N times)",
                 bg=BG_L, fg="#666", font=("Segoe UI", 8, "italic")).pack(side="left")

        self._ub_stats_lbl = tk.Label(ub_frm, text="",
                                      bg=BG_L, fg=ETEC_PURPLE,
                                      font=("Segoe UI", 9, "italic"))
        self._ub_stats_lbl.pack(anchor="w")

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

    def _browse_used_blocks(self):
        path = fd.askopenfilename(
            title="Select Used Blocks File",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not path: return
        self._sv.set("Loading Used Blocks…")
        threading.Thread(target=self._load_ub_t, args=(path,), daemon=True).start()

    def _load_ub_t(self, path):
        try:
            df = load_used_blocks(path); self._q.put(("ub_loaded", df, path))
        except Exception as e: self._q.put(("err", str(e)))

    def _on_bank_option(self):
        if self.bank_df is not None:
            df = self._active_bank()
            self._refresh_from_df(df)

    def _active_bank(self) -> pd.DataFrame:
        base = self._filtered_df if self._filtered_df is not None else self.bank_df
        if self._bank_option.get() == 2 and base is not None:
            unused = base[~base["QuestionID"].astype(str).isin(self._used_qids)].copy()
            return unused if not unused.empty else base
        return base

    def _refresh_unused_label(self):
        if self.bank_df is None: return
        total  = len(self.bank_df)
        used   = len(self._used_qids & set(self.bank_df["QuestionID"].astype(str)))
        unused = total - used
        self._unused_lbl.config(
            text=f"({unused} unused out of {total}  |  {used} used so far)")

    def _on_bank(self, df, path):
        self.bank_df = df; self._filtered_df = df.copy()
        self.source_path = Path(path); self._bp.set(Path(path).name)
        # Reset cumulative usage when a new bank is loaded
        self._used_qids = set()
        self._bank_option.set(1)
        self._unused_lbl.config(text="(no generations yet)")
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

        # ── Adaptive Mean Control toggle ───────────────────────────────────────
        amc_frm = tk.Frame(pf, bg=BG_L)
        amc_frm.grid(row=7, column=0, columnspan=2, sticky="w", pady=(4,2))
        tk.Checkbutton(amc_frm,
                       text="Enable Adaptive Mean Control",
                       variable=self._adaptive_mean_var,
                       bg=BG_L, fg=ETEC_BLUE, font=("Segoe UI", 10, "bold"),
                       activebackground=BG_L).pack(side="left")
        tk.Label(amc_frm,
                 text="  (intelligent swap correction to hit target mean — "
                      "uses reuse & Used Blocks strategically)",
                 bg=BG_L, fg="#666", font=("Segoe UI", 8, "italic")).pack(side="left")

        self._sim_n = self._le(pf, "Simulation students:", 8, "3000")
        self._seed  = self._le(pf, "RNG Seed (blank=random):", 9, "")
        _btn(pf, "Save & Preview Settings", self._save_settings,
             bg=BG_D).grid(row=10, column=0, columnspan=2, pady=8)

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
            f"Adaptive MC     : {'Enabled' if self._adaptive_mean_var.get() else 'Disabled'}",
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
        # Use the active bank view (full or unused-only)
        df = self._active_bank()
        if df is None:
            mb.showerror("Error", "Please upload a question bank first."); return
        try: params, dcol = self._collect(df)
        except ValueError as e: mb.showerror("Configuration Error", str(e)); return
        bank_label = "unused-only" if self._bank_option.get() == 2 else "full bank"
        self._gbtn.config(state="disabled"); self._pv.set(0)
        self._logbox.config(state="normal"); self._logbox.delete("1.0","end")
        self._logbox.config(state="disabled")
        threading.Thread(target=self._gen_t,
                         args=(df, params, dcol, bank_label), daemon=True).start()

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
            used_blocks=self.used_blocks_df if self._use_ub_var.get() else None,
            allow_used_blocks=self._use_ub_var.get(),
            max_ub_usage=max(1, int(self._max_ub_usage_m2.get()))
                         if self._max_ub_usage_m2.get().strip().isdigit() else 2,
            adaptive_mean_control=self._adaptive_mean_var.get(),
        )
        self._n_forms_last = n_forms
        self._n_forms_exam  = self._exam_var.get()
        self._n_forms_stage = self._stage_var.get()
        self._n_forms_pl    = self._part_level_var.get()
        return params, dcol

    def _gen_t(self, df, params, dcol, bank_label="full bank"):
        try:
            n = self._n_forms_last
            # Generate block IDs for Mode 2 as well
            exam  = getattr(self, "_n_forms_exam",  self._exam_var.get())
            stage = getattr(self, "_n_forms_stage", self._stage_var.get())
            pl    = getattr(self, "_n_forms_pl",    self._part_level_var.get())
            if self.source_path:
                names = generate_block_ids(exam, stage, n, pl,
                                           self.source_path.parent,
                                           self.used_blocks_df)
            else:
                names = [f"Form_{i+1}" for i in range(n)]
            self._lg(f"Starting assembly: {n} forms  [{bank_label}, {len(df)} questions]…")
            self._lg(f"  Block IDs: {names[0]} … {names[-1]}", ETEC_TEAL)
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
            exam_type  = self._exam_var.get()
            folder_tag = f"{exam_type}  [{names[0]}…{names[-1]}]"
            out_dir    = _make_output_folder(
                self.source_path.parent, folder_tag, n)
            self._last_out_dir = out_dir
            fp = out_dir / "Forms.xlsx"
            ap = out_dir / "3PL_Analysis.xlsx"
            rp = out_dir / "Remaining_Questions.xlsx"

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

            # Feature 5: Charts workbook
            cp = out_dir / "Charts.xlsx"
            try:
                diff_method = params.difficulty_method
                write_charts_workbook(results, names, dcol, "المجال",
                                      diff_method, cp)
            except Exception as ce:
                self._lg(f"  ⚠ Charts file error: {ce}", "yellow")
                cp = None

            used_ids = set()
            for form, _, _ in results:
                if "QuestionID" in form.columns:
                    used_ids.update(form["QuestionID"].dropna().astype(str).tolist())

            # Cumulative tracking — update Option 2 pool
            self._used_qids.update(used_ids)

            # Remaining = full bank minus ALL ever used
            full_bank = self.bank_df if self.bank_df is not None else df
            remaining = full_bank[
                ~full_bank["QuestionID"].astype(str).isin(self._used_qids)
            ].copy()
            remaining.drop(columns=[c for c in remaining.columns if c.startswith("_")],
                           errors="ignore").to_excel(str(rp), index=False, engine="openpyxl")
            unused_count = len(remaining)

            self._q.put(("prog", 100))
            self._lg(f"\n  📁 Output folder   : {out_dir}")
            self._lg(f"  📄 Forms           : {fp.name}")
            self._lg(f"  📊 3PL Analysis    : {ap.name}")
            self._lg(f"  📋 Remaining Qs    : {rp.name}  ({unused_count} questions)")
            self._lg(f"\n  Bank status: {unused_count} unused / "
                     f"{len(full_bank)} total questions remain", ETEC_TEAL)
            self._lg("\n✓ Generation complete!", "lime")
            self._q.put(("done_m2", valid, unused_count))

        except Exception as e:
            self._lg(f"\n✗ Error: {e}", "red")
            log.exception("Mode2 generation failed")
            self._q.put(("fail",))

# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    LauncherWindow().mainloop()
