"""
Data Loader Module
------------------
Handles loading, validating, and normalising a question bank from Excel/CSV.

Expected columns (Arabic or English aliases accepted):
    QuestionID  : unique identifier
    الناتج      : learning outcome
    المؤشر      : indicator
    المجال      : domain
    b / difficulty : difficulty (IRT b-parameter)
    a / تمييز   : discrimination (IRT a-parameter)
    c / التخمين : guessing (IRT c-parameter)

All three IRT parameters are imputed with sensible defaults when absent.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column alias registry  (canonical → list of acceptable names)
# ---------------------------------------------------------------------------
COLUMN_ALIASES: dict[str, list[str]] = {
    "QuestionID": ["questionid", "question_id", "id", "رقم السؤال", "رقم"],
    "الناتج":     ["الناتج", "outcome", "learning_outcome", "ناتج"],
    "المؤشر":     ["المؤشر", "indicator", "مؤشر"],
    "المجال":     ["المجال", "domain", "مجال"],
    "b":          ["b", "difficulty", "صعوبة", "الصعوبة"],
    "a":          ["a", "تمييز", "discrimination", "التمييز"],
    "c":          ["c", "تخمين", "guessing", "التخمين"],
}

# IRT parameter defaults and valid ranges
IRT_DEFAULTS = {"a": 1.0, "b": 0.0, "c": 0.25}
IRT_RANGES = {
    "a": (0.01, 4.0),
    "b": (-4.0, 4.0),
    "c": (0.0, 0.5),
}


class DataLoadError(Exception):
    """Raised when the question bank cannot be loaded."""


class BankValidator:
    """Validates and cleans a loaded bank DataFrame."""

    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.issues: list[str] = []

    # ------------------------------------------------------------------
    def validate(self) -> pd.DataFrame:
        self._check_required_columns()
        self._deduplicate_ids()
        self._fill_irt_defaults()
        self._clamp_irt_ranges()
        self._create_pair_key()
        if self.issues:
            logger.warning("Bank validation issues:\n" + "\n".join(self.issues))
        return self.df

    # ------------------------------------------------------------------
    def _check_required_columns(self):
        required = {"QuestionID", "الناتج", "المؤشر"}
        missing = required - set(self.df.columns)
        if missing:
            raise DataLoadError(f"Missing required columns after mapping: {missing}")

    def _deduplicate_ids(self):
        dupes = self.df["QuestionID"].duplicated(keep="first")
        n = dupes.sum()
        if n:
            self.issues.append(f"{n} duplicate QuestionID rows removed (kept first).")
            self.df = self.df[~dupes].copy()

    def _fill_irt_defaults(self):
        for param, default in IRT_DEFAULTS.items():
            if param not in self.df.columns:
                self.df[param] = default
                self.issues.append(f"Column '{param}' missing – filled with {default}.")
            else:
                n_null = self.df[param].isna().sum()
                if n_null:
                    self.df[param] = self.df[param].fillna(default)
                    self.issues.append(
                        f"{n_null} missing values in '{param}' filled with {default}."
                    )
                self.df[param] = pd.to_numeric(self.df[param], errors="coerce").fillna(
                    default
                )

    def _clamp_irt_ranges(self):
        for param, (lo, hi) in IRT_RANGES.items():
            out_of_range = ((self.df[param] < lo) | (self.df[param] > hi)).sum()
            if out_of_range:
                self.issues.append(
                    f"{out_of_range} values in '{param}' clamped to [{lo}, {hi}]."
                )
            self.df[param] = self.df[param].clip(lo, hi)

    def _create_pair_key(self):
        """
        Treat (الناتج, المؤشر) as an independent pair.
        The key is a tuple-string so numerically identical المؤشر values
        under different الناتج entries remain distinct.
        """
        self.df["_pair_key"] = (
            self.df["الناتج"].astype(str) + "||" + self.df["المؤشر"].astype(str)
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map raw column names to canonical names using COLUMN_ALIASES."""
    rename_map: dict[str, str] = {}
    lower_cols = {c.strip().lower(): c for c in df.columns}

    for canonical, aliases in COLUMN_ALIASES.items():
        if canonical in df.columns:
            continue  # already present
        for alias in aliases:
            key = alias.strip().lower()
            if key in lower_cols:
                rename_map[lower_cols[key]] = canonical
                break

    return df.rename(columns=rename_map)


def load_bank(path: str | Path) -> pd.DataFrame:
    """
    Load a question bank from an Excel (.xlsx/.xls) or CSV file.

    Returns a validated, normalised DataFrame with columns:
        QuestionID, الناتج, المؤشر, المجال, a, b, c, _pair_key
    and preserves any extra columns present in the source file.
    """
    path = Path(path)
    if not path.exists():
        raise DataLoadError(f"File not found: {path}")

    suffix = path.suffix.lower()
    try:
        if suffix in (".xlsx", ".xls", ".xlsm"):
            raw = pd.read_excel(path, engine="openpyxl")
        elif suffix == ".csv":
            raw = pd.read_csv(path, encoding="utf-8-sig")
        else:
            raise DataLoadError(f"Unsupported file type: {suffix}")
    except Exception as exc:
        raise DataLoadError(f"Could not read file: {exc}") from exc

    if raw.empty:
        raise DataLoadError("The file is empty.")

    df = _normalise_columns(raw)
    df = BankValidator(df).validate()

    logger.info(
        "Loaded %d questions from '%s'. Columns: %s",
        len(df),
        path.name,
        list(df.columns),
    )
    return df.reset_index(drop=True)


def get_pair_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a summary DataFrame of (الناتج, المؤشر) pairs with
    question counts and mean/min/max difficulty.
    """
    grp = df.groupby("_pair_key", sort=False)
    summary = grp.agg(
        الناتج=("الناتج", "first"),
        المؤشر=("المؤشر", "first"),
        count=("QuestionID", "count"),
        mean_b=("b", "mean"),
        min_b=("b", "min"),
        max_b=("b", "max"),
    ).reset_index(drop=True)
    return summary


def get_domain_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return question counts per domain."""
    if "المجال" not in df.columns:
        return pd.DataFrame()
    return (
        df.groupby("المجال", sort=False)
        .agg(count=("QuestionID", "count"), mean_b=("b", "mean"))
        .reset_index()
    )
