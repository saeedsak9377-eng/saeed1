"""
Smart Test Assembly & Psychometric Analysis
===========================================
Single-file version — copy this ONE file to your machine and run:

    python test_assembly_all_in_one.py

Requirements (install once):
    pip install pandas numpy openpyxl matplotlib xlsxwriter
"""

# ============================================================
# STANDARD LIBRARY
# ============================================================
from __future__ import annotations

import argparse
import io
import logging
import math
import queue
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

# ============================================================
# THIRD-PARTY (no scipy)
# ============================================================
import matplotlib
matplotlib.use("Agg")
import matplotlib.cm as cm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlsxwriter

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-8s  %(message)s")
logger = logging.getLogger("test_assembly")


# ════════════════════════════════════════════════════════════
#  MODULE 1 — DATA LOADER
# ════════════════════════════════════════════════════════════

COLUMN_ALIASES: dict[str, list[str]] = {
    "QuestionID": ["questionid", "question_id", "id", "رقم السؤال", "رقم"],
    "الناتج":     ["الناتج", "outcome", "learning_outcome", "ناتج"],
    "المؤشر":     ["المؤشر", "indicator", "مؤشر"],
    "المجال":     ["المجال", "domain", "مجال"],
    "b":          ["b", "difficulty", "صعوبة", "الصعوبة"],
    "a":          ["a", "تمييز", "discrimination", "التمييز"],
    "c":          ["c", "تخمين", "guessing", "التخمين"],
}

IRT_DEFAULTS = {"a": 1.0, "b": 0.0, "c": 0.25}
IRT_RANGES   = {"a": (0.01, 4.0), "b": (-4.0, 4.0), "c": (0.0, 0.5)}


class DataLoadError(Exception):
    pass


class BankValidator:
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.issues: list[str] = []

    def validate(self) -> pd.DataFrame:
        self._check_required_columns()
        self._deduplicate_ids()
        self._fill_irt_defaults()
        self._clamp_irt_ranges()
        self._create_pair_key()
        if self.issues:
            logger.warning("Bank issues:\n" + "\n".join(self.issues))
        return self.df

    def _check_required_columns(self):
        missing = {"QuestionID", "الناتج", "المؤشر"} - set(self.df.columns)
        if missing:
            raise DataLoadError(f"Missing required columns: {missing}")

    def _deduplicate_ids(self):
        dupes = self.df["QuestionID"].duplicated(keep="first")
        n = dupes.sum()
        if n:
            self.issues.append(f"{n} duplicate QuestionID rows removed.")
            self.df = self.df[~dupes].copy()

    def _fill_irt_defaults(self):
        for param, default in IRT_DEFAULTS.items():
            if param not in self.df.columns:
                self.df[param] = default
                self.issues.append(f"'{param}' missing — filled with {default}.")
            else:
                self.df[param] = pd.to_numeric(self.df[param],
                                               errors="coerce").fillna(default)

    def _clamp_irt_ranges(self):
        for param, (lo, hi) in IRT_RANGES.items():
            n_bad = ((self.df[param] < lo) | (self.df[param] > hi)).sum()
            if n_bad:
                self.issues.append(f"{n_bad} values in '{param}' clamped to [{lo},{hi}].")
            self.df[param] = self.df[param].clip(lo, hi)

    def _create_pair_key(self):
        self.df["_pair_key"] = (
            self.df["الناتج"].astype(str) + "||" + self.df["المؤشر"].astype(str)
        )


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    lower_cols = {c.strip().lower(): c for c in df.columns}
    for canonical, aliases in COLUMN_ALIASES.items():
        if canonical in df.columns:
            continue
        for alias in aliases:
            key = alias.strip().lower()
            if key in lower_cols:
                rename_map[lower_cols[key]] = canonical
                break
    return df.rename(columns=rename_map)


def load_bank(path: str | Path) -> pd.DataFrame:
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
    logger.info("Loaded %d questions from '%s'", len(df), path.name)
    return df.reset_index(drop=True)


def get_pair_summary(df: pd.DataFrame) -> pd.DataFrame:
    grp = df.groupby("_pair_key", sort=False)
    return grp.agg(
        الناتج=("الناتج", "first"),
        المؤشر=("المؤشر", "first"),
        count=("QuestionID", "count"),
        mean_b=("b", "mean"),
        min_b=("b", "min"),
        max_b=("b", "max"),
    ).reset_index(drop=True)


# ════════════════════════════════════════════════════════════
#  MODULE 2 — ASSEMBLY ENGINE
# ════════════════════════════════════════════════════════════

@dataclass
class BinSpec:
    low: float
    high: float
    count: int


@dataclass
class PairSpec:
    pair_key: str
    الناتج: str
    المؤشر: str
    count: int


@dataclass
class AssemblyConfig:
    n_forms: int = 1
    questions_per_form: int = 40
    target_mean_b: float = 0.0
    target_min_b: float = -3.0
    target_max_b: float = 3.0
    bins: list[tuple[float, float, float]] = field(default_factory=list)
    pair_requirements: list[dict] = field(default_factory=list)
    w_difficulty: float = 0.40
    w_bin: float        = 0.25
    w_usage: float      = 0.20
    w_random: float     = 0.15
    allow_reuse_across_forms: bool = False
    rng_seed: Optional[int] = None


@dataclass
class AssemblyResult:
    forms: list[pd.DataFrame]
    stats: list[dict]
    warnings: list[str]
    usage_counts: dict[str, int]


def _largest_remainder(total: int, proportions: list[float]) -> list[int]:
    if not proportions:
        return []
    s = sum(proportions)
    floats = [total * p / s for p in proportions]
    floors = [int(f) for f in floats]
    remainders = sorted(enumerate(floats), key=lambda x: -(x[1] - int(x[1])))
    deficit = total - sum(floors)
    for k in range(deficit):
        floors[remainders[k][0]] += 1
    return floors


def _irt_b_score(b: np.ndarray, target: float, b_range: tuple) -> np.ndarray:
    span = max(b_range[1] - b_range[0], 1e-6)
    return 1.0 - np.clip(np.abs(b - target) / span, 0.0, 1.0)


def _bin_membership_score(b: np.ndarray, lo: float, hi: float) -> np.ndarray:
    mid = (lo + hi) / 2.0
    hw  = max((hi - lo) / 2.0, 1e-6)
    return np.clip(1.0 - (np.abs(b - mid) - hw) / hw, 0.0, 1.0)


class AssemblyEngine:
    def __init__(self, bank: pd.DataFrame, config: AssemblyConfig):
        self.bank   = bank.copy().reset_index(drop=True)
        self.config = config
        self.rng    = np.random.default_rng(config.rng_seed)
        self.usage_counts: dict[str, int] = {
            str(q): 0 for q in self.bank["QuestionID"]
        }
        self.warnings: list[str] = []

    def assemble(self) -> AssemblyResult:
        cfg   = self.config
        forms = []
        stats = []
        bins  = self._resolve_bins()
        pairs = self._resolve_pairs()
        for idx in range(cfg.n_forms):
            logger.info("Assembling form %d / %d", idx + 1, cfg.n_forms)
            form_df, form_stats = self._assemble_one(idx, bins, pairs)
            forms.append(form_df)
            stats.append(form_stats)
        return AssemblyResult(forms=forms, stats=stats,
                              warnings=self.warnings,
                              usage_counts=dict(self.usage_counts))

    # ---- bin & pair resolution ----
    def _resolve_bins(self) -> list[BinSpec]:
        cfg = self.config
        n   = cfg.questions_per_form
        if not cfg.bins:
            return [BinSpec(cfg.target_min_b, cfg.target_max_b, n)]
        proportions = [p for _, _, p in cfg.bins]
        counts = _largest_remainder(n, proportions)
        result = [BinSpec(lo, hi, cnt) for (lo, hi, _), cnt in zip(cfg.bins, counts)]
        diff = n - sum(b.count for b in result)
        if diff:
            result[-1].count += diff
        return result

    def _resolve_pairs(self) -> list[PairSpec]:
        cfg = self.config
        if not cfg.pair_requirements:
            return []
        total = sum(r["count"] for r in cfg.pair_requirements)
        if total != cfg.questions_per_form:
            counts = _largest_remainder(cfg.questions_per_form,
                                        [r["count"] for r in cfg.pair_requirements])
        else:
            counts = [r["count"] for r in cfg.pair_requirements]
        return [PairSpec(r["pair_key"], r.get("الناتج",""),
                         r.get("المؤشر",""), cnt)
                for r, cnt in zip(cfg.pair_requirements, counts)]

    # ---- main assembly ----
    def _assemble_one(self, idx, bins, pairs):
        selected_ids: set[str] = set()
        rows: list[pd.DataFrame] = []
        if pairs:
            rows, selected_ids = self._by_pairs_bins(pairs, bins, idx)
        else:
            rows, selected_ids = self._by_bins(bins, idx)
        form_df = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        for qid in selected_ids:
            self.usage_counts[str(qid)] = self.usage_counts.get(str(qid), 0) + 1
        return form_df, self._form_stats(form_df)

    def _by_pairs_bins(self, pairs, bins, idx):
        selected_ids: set[str] = set()
        all_rows: list[pd.DataFrame] = []
        matrix = self._pair_bin_matrix(pairs, bins)
        for pi, pair in enumerate(pairs):
            pool = self._pool(pair.pair_key, selected_ids)
            for bi, bin_spec in enumerate(bins):
                needed = matrix[pi][bi]
                if needed == 0:
                    continue
                r, ids = self._select(pool, needed, bin_spec, selected_ids)
                all_rows.extend(r); selected_ids.update(ids)
                pool = pool[~pool["QuestionID"].isin(selected_ids)]
                shortage = needed - len(ids)
                if shortage > 0:
                    self.warnings.append(
                        f"Form {idx+1}, pair '{pair.pair_key}', "
                        f"bin [{bin_spec.low:.2f},{bin_spec.high:.2f}]: "
                        f"shortage {shortage} — fallback applied.")
                    fr, fi = self._fallback(shortage, bin_spec, selected_ids)
                    all_rows.extend(fr); selected_ids.update(fi)
        return all_rows, selected_ids

    def _by_bins(self, bins, idx):
        selected_ids: set[str] = set()
        all_rows: list[pd.DataFrame] = []
        for bin_spec in bins:
            pool = self._pool(None, selected_ids)
            r, ids = self._select(pool, bin_spec.count, bin_spec, selected_ids)
            all_rows.extend(r); selected_ids.update(ids)
            shortage = bin_spec.count - len(ids)
            if shortage > 0:
                self.warnings.append(
                    f"Form {idx+1}, bin [{bin_spec.low:.2f},{bin_spec.high:.2f}]: "
                    f"shortage {shortage} — fallback applied.")
                fr, fi = self._fallback(shortage, bin_spec, selected_ids)
                all_rows.extend(fr); selected_ids.update(fi)
        return all_rows, selected_ids

    # ---- pool & scoring ----
    def _pool(self, pair_key: Optional[str], exclude: set[str]) -> pd.DataFrame:
        pool = self.bank.copy()
        if pair_key is not None:
            pool = pool[pool["_pair_key"] == pair_key]
        if exclude:
            pool = pool[~pool["QuestionID"].isin(exclude)]
        if not self.config.allow_reuse_across_forms:
            used = {q for q, c in self.usage_counts.items() if c > 0}
            pool = pool[~pool["QuestionID"].isin(used)]
        return pool.reset_index(drop=True)

    def _score(self, pool: pd.DataFrame, bin_spec: BinSpec) -> np.ndarray:
        cfg = self.config
        b    = pool["b"].to_numpy(float)
        qids = pool["QuestionID"].astype(str).tolist()
        diff  = _irt_b_score(b, cfg.target_mean_b,
                              (cfg.target_min_b, cfg.target_max_b))
        bfit  = _bin_membership_score(b, bin_spec.low, bin_spec.high)
        usage = np.array([self.usage_counts.get(q, 0) for q in qids], float)
        mu    = max(usage.max(), 1.0)
        noise = self.rng.random(len(pool))
        return (cfg.w_difficulty * diff + cfg.w_bin * bfit
                + cfg.w_usage * (1.0 - usage / mu)
                + cfg.w_random * noise)

    def _select(self, pool, needed, bin_spec, exclude):
        if pool.empty or needed == 0:
            return [], set()
        cands = pool[(pool["b"] >= bin_spec.low) &
                     (pool["b"] <= bin_spec.high)].copy()
        if len(cands) < needed:
            hw = (bin_spec.high - bin_spec.low) / 2.0
            cands = pool[(pool["b"] >= bin_spec.low - 0.2*hw) &
                         (pool["b"] <= bin_spec.high + 0.2*hw)].copy()
        if cands.empty:
            cands = pool.copy()
        scores = self._score(cands, bin_spec)
        top = cands.iloc[np.argsort(-scores)[:needed]]
        return [top], set(top["QuestionID"].astype(str).tolist())

    def _fallback(self, shortage, bin_spec, exclude):
        pool = self.bank[~self.bank["QuestionID"].isin(exclude)].copy()
        if not self.config.allow_reuse_across_forms:
            used = {q for q, c in self.usage_counts.items() if c > 0}
            pool = pool[~pool["QuestionID"].isin(used)]
        if pool.empty:
            pool = self.bank[~self.bank["QuestionID"].isin(exclude)].copy()
        if pool.empty:
            return [], set()
        mid  = (bin_spec.low + bin_spec.high) / 2.0
        dist = np.abs(pool["b"].to_numpy(float) - mid)
        usag = np.array([self.usage_counts.get(str(q), 0)
                         for q in pool["QuestionID"]], float)
        top  = pool.iloc[np.lexsort((dist, usag))[:shortage]]
        return [top], set(top["QuestionID"].astype(str).tolist())

    def _pair_bin_matrix(self, pairs, bins):
        matrix = []
        for pair in pairs:
            pb = self.bank[self.bank["_pair_key"] == pair.pair_key]
            avail = [max(((pb["b"] >= b.low) & (pb["b"] <= b.high)).sum(), 1)
                     for b in bins]
            matrix.append(_largest_remainder(pair.count, avail))
        return matrix

    def _form_stats(self, df: pd.DataFrame) -> dict:
        if df.empty:
            return {}
        return {
            "n_items": len(df),
            "mean_b": float(df["b"].mean()),
            "std_b":  float(df["b"].std()),
            "min_b":  float(df["b"].min()),
            "max_b":  float(df["b"].max()),
            "mean_a": float(df["a"].mean()),
            "mean_c": float(df["c"].mean()),
        }


def build_config_from_gui(params: dict) -> AssemblyConfig:
    bins = [(b["low"], b["high"], b["proportion"])
            for b in params.get("bins", [])]
    return AssemblyConfig(
        n_forms=int(params.get("n_forms", 1)),
        questions_per_form=int(params.get("questions_per_form", 40)),
        target_mean_b=float(params.get("target_mean_b", 0.0)),
        target_min_b=float(params.get("target_min_b", -3.0)),
        target_max_b=float(params.get("target_max_b", 3.0)),
        bins=bins,
        pair_requirements=params.get("pair_requirements", []),
        w_difficulty=float(params.get("w_difficulty", 0.40)),
        w_bin=float(params.get("w_bin", 0.25)),
        w_usage=float(params.get("w_usage", 0.20)),
        w_random=float(params.get("w_random", 0.15)),
        allow_reuse_across_forms=bool(params.get("allow_reuse_across_forms", False)),
        rng_seed=params.get("rng_seed"),
    )


# ════════════════════════════════════════════════════════════
#  MODULE 3 — PSYCHOMETRICS (3PL, no scipy)
# ════════════════════════════════════════════════════════════

def p3pl(theta, a, b, c) -> np.ndarray:
    theta = np.asarray(theta, float).reshape(-1, 1)
    a = np.asarray(a, float).reshape(1, -1)
    b = np.asarray(b, float).reshape(1, -1)
    c = np.asarray(c, float).reshape(1, -1)
    return c + (1.0 - c) / (1.0 + np.exp(-a * (theta - b)))


def item_information(theta, a, b, c) -> np.ndarray:
    P = p3pl(theta, a, b, c)
    Q = 1.0 - P
    a2 = np.asarray(a, float).reshape(1, -1)
    c2 = np.asarray(c, float).reshape(1, -1)
    num = ((P - c2) / (1.0 - c2 + 1e-12)) ** 2
    return a2**2 * num * (Q / (P + 1e-12))


@dataclass
class SimulationConfig:
    n_students: int = 3000
    theta_mean: float = 0.0
    theta_sd: float = 1.0
    theta_distribution: str = "normal"
    rng_seed: Optional[int] = None


def simulate_responses(form_df, sim_cfg=None):
    if sim_cfg is None:
        sim_cfg = SimulationConfig()
    rng = np.random.default_rng(sim_cfg.rng_seed)
    N   = sim_cfg.n_students
    if sim_cfg.theta_distribution == "uniform":
        theta = rng.uniform(-3.0, 3.0, size=N)
    else:
        theta = rng.normal(sim_cfg.theta_mean, sim_cfg.theta_sd, size=N)
    a, b, c = (form_df[k].to_numpy(float) for k in ("a", "b", "c"))
    P = p3pl(theta, a, b, c)
    X = (rng.random((N, len(form_df))) < P).astype(np.int8)
    return theta, X


def cronbach_alpha(X: np.ndarray) -> float:
    N, K = X.shape
    if K < 2:
        return float("nan")
    total_var = X.sum(axis=1).var(ddof=1)
    if total_var == 0:
        return float("nan")
    alpha = (K / (K - 1)) * (1.0 - X.var(axis=0, ddof=1).sum() / total_var)
    return float(np.clip(alpha, -1.0, 1.0))


def item_total_correlations(X: np.ndarray) -> np.ndarray:
    K = X.shape[1]
    cors = np.zeros(K)
    total = X.sum(axis=1).astype(float)
    for j in range(K):
        rest = total - X[:, j]
        if rest.std() == 0 or X[:, j].std() == 0:
            cors[j] = float("nan")
        else:
            cors[j] = np.corrcoef(X[:, j].astype(float), rest)[0, 1]
    return cors


@dataclass
class FormAnalysis:
    form_name: str
    form_df: pd.DataFrame
    theta_grid: np.ndarray
    P_grid: np.ndarray
    TCC: np.ndarray
    TIF: np.ndarray
    SEM: np.ndarray
    alpha: float
    item_cors: np.ndarray
    sim_theta: np.ndarray
    sim_X: np.ndarray
    summary: dict


def analyse_form(form_df, form_name="Form", sim_cfg=None) -> FormAnalysis:
    if form_df.empty:
        raise ValueError(f"Form '{form_name}' is empty.")
    a, b, c = (form_df[k].to_numpy(float) for k in ("a", "b", "c"))
    theta_grid = np.linspace(-4.0, 4.0, 200)
    P_grid  = p3pl(theta_grid, a, b, c)
    TCC     = P_grid.sum(axis=1)
    info    = item_information(theta_grid, a, b, c)
    TIF     = info.sum(axis=1)
    SEM     = 1.0 / np.sqrt(np.clip(TIF, 1e-8, None))
    th_sim, X_sim = simulate_responses(form_df, sim_cfg)
    alpha   = cronbach_alpha(X_sim)
    cors    = item_total_correlations(X_sim)
    summary = {
        "n_items":      len(form_df),
        "mean_b":       float(np.mean(b)),
        "std_b":        float(np.std(b)),
        "min_b":        float(np.min(b)),
        "max_b":        float(np.max(b)),
        "mean_a":       float(np.mean(a)),
        "mean_c":       float(np.mean(c)),
        "alpha":        alpha,
        "mean_item_cor": float(np.nanmean(cors)),
        "min_item_cor":  float(np.nanmin(cors)),
        "max_item_cor":  float(np.nanmax(cors)),
        "peak_info_theta": float(theta_grid[np.argmax(TIF)]),
        "peak_info":       float(np.max(TIF)),
    }
    return FormAnalysis(form_name=form_name, form_df=form_df,
                        theta_grid=theta_grid, P_grid=P_grid,
                        TCC=TCC, TIF=TIF, SEM=SEM,
                        alpha=alpha, item_cors=cors,
                        sim_theta=th_sim, sim_X=X_sim, summary=summary)


def analyse_all_forms(forms, form_names=None, sim_cfg=None):
    if form_names is None:
        form_names = [f"Form {i+1}" for i in range(len(forms))]
    out = []
    for df, name in zip(forms, form_names):
        try:
            out.append(analyse_form(df, name, sim_cfg))
        except Exception as exc:
            logger.error("Analysis failed for '%s': %s", name, exc)
    return out


# ════════════════════════════════════════════════════════════
#  MODULE 4 — OUTPUT GENERATOR
# ════════════════════════════════════════════════════════════

HDR_BG   = "#1F3864"
HDR_FG   = "#FFFFFF"
SUB_BG   = "#2E75B6"
ALT_BG   = "#DCE6F1"
CHART_CL = ["#2E75B6","#ED7D31","#A9D18E","#FF0000",
             "#7030A0","#00B0F0","#FF6600","#70AD47"]

FORM_COLS = [
    ("QuestionID", 18, False), ("الناتج", 22, False),
    ("المؤشر", 22, False),    ("المجال", 20, False),
    ("b", 12, True),           ("a", 12, True),
    ("c", 12, True),
]


def _fmts(wb):
    f = {}
    f["hdr"]  = wb.add_format({"bold":True,"bg_color":HDR_BG,"font_color":HDR_FG,
                                 "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    f["sub"]  = wb.add_format({"bold":True,"bg_color":SUB_BG,"font_color":HDR_FG,
                                 "border":1,"align":"center"})
    f["mlbl"] = wb.add_format({"bold":True,"bg_color":"#BDD7EE","border":1})
    f["mval"] = wb.add_format({"bg_color":"#FFFFFF","border":1,"num_format":"0.0000"})
    f["mvi"]  = wb.add_format({"bg_color":"#FFFFFF","border":1,"num_format":"0"})
    f["dat"]  = wb.add_format({"border":1,"align":"left","valign":"vcenter"})
    f["num"]  = wb.add_format({"border":1,"align":"center","num_format":"0.0000"})
    f["dat2"] = wb.add_format({"bg_color":ALT_BG,"border":1,"align":"left"})
    f["num2"] = wb.add_format({"bg_color":ALT_BG,"border":1,"align":"center","num_format":"0.0000"})
    f["ttl"]  = wb.add_format({"bold":True,"font_size":14,"bg_color":HDR_BG,
                                "font_color":HDR_FG,"align":"center","valign":"vcenter"})
    return f


def _to_png(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


# ---- forms workbook ----
def write_forms_workbook(analyses, path):
    wb = xlsxwriter.Workbook(str(path))
    f  = _fmts(wb)
    for fa in analyses:
        ws = wb.add_worksheet(fa.form_name[:31])
        ws.right_to_left()
        s  = fa.summary
        ws.merge_range("A1:D1", fa.form_name, f["ttl"])
        metrics = [("N Items", s.get("n_items"), False),
                   ("Mean b",  s.get("mean_b"),  True),
                   ("Min b",   s.get("min_b"),   True),
                   ("Max b",   s.get("max_b"),   True),
                   ("SD b",    s.get("std_b"),   True),
                   ("Mean a",  s.get("mean_a"),  True),
                   ("Mean c",  s.get("mean_c"),  True),
                   ("Alpha",   s.get("alpha"),   True),
                   ("Mean Corr.", s.get("mean_item_cor"), True)]
        ws.write(2, 0, "Metric", f["sub"]); ws.write(2, 1, "Value", f["sub"])
        for ri, (lbl, val, is_f) in enumerate(metrics):
            r = 3 + ri
            ws.write(r, 0, lbl, f["mlbl"])
            fmt = f["mval"] if is_f else f["mvi"]
            if val is None:
                ws.write(r, 1, "N/A", f["dat"])
            elif is_f:
                ws.write_number(r, 1, float(val), fmt)
            else:
                ws.write_number(r, 1, int(val), fmt)
        ws.set_column(0, 0, 30); ws.set_column(1, 1, 15)

        df = fa.form_df.copy().reset_index(drop=True)
        if fa.item_cors is not None and len(fa.item_cors) == len(df):
            df["Item Corr."] = np.round(fa.item_cors, 4)
        cols = list(FORM_COLS)
        if "Item Corr." in df.columns:
            cols.append(("Item Corr.", 14, True))
        start = 13
        for ci, (cn, w, _) in enumerate(cols):
            ws.write(start, ci, cn, f["sub"]); ws.set_column(ci, ci, w)
        for ri, (_, row) in enumerate(df.iterrows()):
            alt = ri % 2 == 1
            for ci, (cn, _, is_n) in enumerate(cols):
                val = row.get(cn, "")
                if is_n:
                    fmt = f["num2"] if alt else f["num"]
                    try:
                        ws.write_number(start+1+ri, ci, float(val), fmt)
                    except Exception:
                        ws.write(start+1+ri, ci, str(val), fmt)
                else:
                    fmt = f["dat2"] if alt else f["dat"]
                    ws.write(start+1+ri, ci, str(val) if val is not None else "", fmt)
        ws.freeze_panes(start+1, 0)
        ws.autofilter(start, 0, start+len(df), len(cols)-1)
    wb.close()
    logger.info("Forms → %s", path)


# ---- global summary workbook ----
def write_global_summary(analyses, path):
    wb = xlsxwriter.Workbook(str(path))
    f  = _fmts(wb)
    ws = wb.add_worksheet("Cross-Form Summary")
    cols = ["Form","n_items","mean_b","std_b","min_b","max_b",
            "mean_a","mean_c","alpha","mean_item_cor",
            "min_item_cor","max_item_cor","peak_info_theta","peak_info"]
    hdrs = {"Form":"Form","n_items":"N Items","mean_b":"Mean b","std_b":"SD b",
            "min_b":"Min b","max_b":"Max b","mean_a":"Mean a","mean_c":"Mean c",
            "alpha":"Alpha","mean_item_cor":"Mean Corr.","min_item_cor":"Min Corr.",
            "max_item_cor":"Max Corr.","peak_info_theta":"Peak θ","peak_info":"Peak Info"}
    ws.merge_range(0,0,0,len(cols)-1,"Cross-Form Psychometric Summary",f["ttl"])
    ws.set_row(0,22)
    for ci,col in enumerate(cols):
        ws.write(1,ci,hdrs[col],f["sub"])
        ws.set_column(ci,ci,14)
    for ri,fa in enumerate(analyses):
        s=fa.summary; alt=ri%2==0
        row=[fa.form_name,s.get("n_items"),s.get("mean_b"),s.get("std_b"),
             s.get("min_b"),s.get("max_b"),s.get("mean_a"),s.get("mean_c"),
             s.get("alpha"),s.get("mean_item_cor"),s.get("min_item_cor"),
             s.get("max_item_cor"),s.get("peak_info_theta"),s.get("peak_info")]
        for ci,val in enumerate(row):
            df_=f["dat2"] if alt else f["dat"]
            dn_=f["num2"] if alt else f["num"]
            if ci==0: ws.write(2+ri,ci,str(val),df_)
            elif val is None: ws.write(2+ri,ci,"N/A",df_)
            elif ci==1: ws.write_number(2+ri,ci,int(val),df_)
            else: ws.write_number(2+ri,ci,float(val),dn_)
    ws.freeze_panes(2,0)
    wb.close()
    logger.info("Summary → %s", path)


# ---- psychometric workbook ----
def write_psychometric_workbook(analyses, path):
    wb = xlsxwriter.Workbook(str(path))
    f  = _fmts(wb)
    for fa in analyses:
        ws = wb.add_worksheet(f"ICC - {fa.form_name}"[:31])
        ws.merge_range("A1:B1", f"ICC – {fa.form_name}", f["ttl"])
        fig, ax = plt.subplots(figsize=(10,6))
        K = fa.P_grid.shape[1]
        cmap = cm.get_cmap("tab20", K)
        for j in range(K):
            qid = str(fa.form_df["QuestionID"].iloc[j]) if j < len(fa.form_df) else f"Item{j+1}"
            ax.plot(fa.theta_grid, fa.P_grid[:,j],
                    color=cmap(j), lw=1.2, alpha=0.85, label=qid)
        ax.set(xlabel="Ability (θ)", ylabel="P(correct)",
               title=f"ICC – {fa.form_name}", xlim=(-4,4), ylim=(0,1.05))
        ax.grid(alpha=0.3)
        if K<=20: ax.legend(loc="upper left",fontsize=7,ncol=2,framealpha=0.7)
        fig.tight_layout()
        ws.insert_image("D2","",{"image_data":_to_png(fig)})
        plt.close(fig)

    ws = wb.add_worksheet("TCC")
    ws.merge_range("A1:B1","Test Characteristic Curves",f["ttl"])
    fig,ax = plt.subplots(figsize=(10,5))
    for i,fa in enumerate(analyses):
        ax.plot(fa.theta_grid,fa.TCC,color=CHART_CL[i%len(CHART_CL)],lw=2,label=fa.form_name)
    ax.set(xlabel="Ability (θ)",ylabel="Expected Score",
           title="TCC",xlim=(-4,4)); ax.grid(alpha=0.3); ax.legend()
    fig.tight_layout(); ws.insert_image("D2","",{"image_data":_to_png(fig)}); plt.close(fig)

    ws = wb.add_worksheet("TIF")
    ws.merge_range("A1:B1","Test Information Functions",f["ttl"])
    fig,axes = plt.subplots(1,2,figsize=(14,5))
    for i,fa in enumerate(analyses):
        c=CHART_CL[i%len(CHART_CL)]
        axes[0].plot(fa.theta_grid,fa.TIF,color=c,lw=2,label=fa.form_name)
        axes[1].plot(fa.theta_grid,fa.SEM,color=c,lw=2,label=fa.form_name)
    axes[0].set(xlabel="θ",ylabel="Information",title="TIF",xlim=(-4,4)); axes[0].grid(alpha=0.3); axes[0].legend()
    axes[1].set(xlabel="θ",ylabel="SEM",title="SEM",xlim=(-4,4),ylim=(0,2)); axes[1].grid(alpha=0.3); axes[1].legend()
    fig.tight_layout(); ws.insert_image("D2","",{"image_data":_to_png(fig)}); plt.close(fig)

    ws = wb.add_worksheet("Item Correlations")
    ws.merge_range("A1:C1","Corrected Item-Total Correlations",f["ttl"])
    row=2
    for fa in analyses:
        ws.write(row,0,fa.form_name,f["sub"]); ws.write(row,1,"QuestionID",f["sub"])
        ws.write(row,2,"Item-Total Corr.",f["sub"]); ws.write(row,3,"b",f["sub"])
        row+=1
        for j in range(len(fa.form_df)):
            qid=str(fa.form_df["QuestionID"].iloc[j])
            bv=float(fa.form_df["b"].iloc[j])
            cor=float(fa.item_cors[j]) if not np.isnan(fa.item_cors[j]) else 0.0
            alt=j%2==0
            df_=f["dat2"] if alt else f["dat"]; dn_=f["num2"] if alt else f["num"]
            ws.write(row+j,0,str(j+1),df_); ws.write(row+j,1,qid,df_)
            ws.write_number(row+j,2,cor,dn_); ws.write_number(row+j,3,bv,dn_)
        row+=len(fa.form_df)+1
    ws.set_column(0,0,8); ws.set_column(1,1,20); ws.set_column(2,2,18); ws.set_column(3,3,15)

    fig,ax=plt.subplots(figsize=(8,4))
    data=[fa.item_cors[~np.isnan(fa.item_cors)] for fa in analyses if fa.item_cors is not None]
    labels=[fa.form_name for fa in analyses if fa.item_cors is not None]
    if data:
        bp=ax.boxplot(data,labels=labels,patch_artist=True)
        for patch,col in zip(bp["boxes"],CHART_CL):
            patch.set_facecolor(col); patch.set_alpha(0.7)
    ax.set(ylabel="Item-Total Corr.",title="Correlation Distribution by Form"); ax.grid(alpha=0.3,axis="y")
    fig.tight_layout(); ws.insert_image(2,5,"",{"image_data":_to_png(fig)}); plt.close(fig)
    wb.close()
    logger.info("Psychometric → %s", path)


def generate_all_outputs(analyses, output_dir: Path, prefix="test"):
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "forms":        output_dir / f"{prefix}_forms.xlsx",
        "summary":      output_dir / f"{prefix}_global_summary.xlsx",
        "psychometric": output_dir / f"{prefix}_psychometric.xlsx",
    }
    write_forms_workbook(analyses, paths["forms"])
    write_global_summary(analyses, paths["summary"])
    write_psychometric_workbook(analyses, paths["psychometric"])
    return paths


# ════════════════════════════════════════════════════════════
#  MODULE 5 — GUI
# ════════════════════════════════════════════════════════════

BG_DARK   = "#1F3864"
BG_MID    = "#2E75B6"
BG_LIGHT  = "#EBF3FB"
ACCENT    = "#ED7D31"
TEXT_DARK = "#1A1A2E"
TEXT_LT   = "#FFFFFF"
BTN_ACT   = "#C55A11"

FT = ("Segoe UI", 15, "bold")
FH = ("Segoe UI", 11, "bold")
FB = ("Segoe UI", 10)
FM = ("Consolas", 9)


def _e(parent, w=14, **kw):
    return tk.Entry(parent, width=w, bg="#FFFFFF", fg="#000000",
                    font=FB, relief="solid", bd=1, **kw)


def _lbl(parent, text, bold=False, **kw):
    return tk.Label(parent, text=text, bg=parent.cget("bg"),
                    fg=TEXT_DARK, font=(FH if bold else FB), **kw)


def _btn(parent, text, cmd, bg=BG_MID, **kw):
    return tk.Button(parent, text=text, command=cmd,
                     bg=bg, fg=TEXT_LT, font=FB,
                     activebackground=BTN_ACT, activeforeground=TEXT_LT,
                     relief="flat", padx=10, pady=5, cursor="hand2", **kw)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Test Assembly & Psychometric Analysis")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)
        self.minsize(920, 680)
        self.bank_df  = None
        self.analyses = None
        self.output_paths = {}
        self._q: queue.Queue = queue.Queue()
        self._build_header()
        self._build_nb()
        self._build_status()
        self.after(80, self._poll)

    # ── header ──
    def _build_header(self):
        frm = tk.Frame(self, bg=BG_DARK, pady=8)
        frm.pack(fill="x")
        tk.Label(frm, text="Smart Test Assembly  &  Psychometric Analysis",
                 bg=BG_DARK, fg=TEXT_LT, font=FT).pack()

    # ── notebook ──
    def _build_nb(self):
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("TNotebook", background=BG_DARK, borderwidth=0)
        s.configure("TNotebook.Tab", background=BG_MID, foreground=TEXT_LT,
                    font=FH, padding=[14,6])
        s.map("TNotebook.Tab", background=[("selected",ACCENT)],
              foreground=[("selected",TEXT_LT)])
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(0,4))
        t1=self._scroll_tab(); t2=self._scroll_tab()
        t3=tk.Frame(self.nb, bg=BG_LIGHT)
        t4=self._scroll_tab()
        self.nb.add(t1, text="  1. Bank  ")
        self.nb.add(t2, text="  2. Configure  ")
        self.nb.add(t3, text="  3. Generate  ")
        self.nb.add(t4, text="  4. Results  ")
        self._build_bank(t1._inner)
        self._build_config(t2._inner)
        self._build_generate(t3)
        self._build_results(t4._inner)

    def _scroll_tab(self):
        outer  = tk.Frame(self.nb, bg=BG_LIGHT)
        canvas = tk.Canvas(outer, bg=BG_LIGHT, highlightthickness=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner  = tk.Frame(canvas, bg=BG_LIGHT)
        wid    = canvas.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
        outer._inner = inner
        return outer

    # ── status bar ──
    def _build_status(self):
        frm = tk.Frame(self, bg=BG_DARK, height=24)
        frm.pack(fill="x", side="bottom")
        self._sv = tk.StringVar(value="Ready")
        tk.Label(frm, textvariable=self._sv, bg=BG_DARK, fg=TEXT_LT,
                 font=FM, anchor="w", padx=8).pack(fill="x")

    def _status(self, msg):
        self._sv.set(msg)

    # ══════════════ TAB 1 — BANK ══════════════
    def _build_bank(self, p):
        p.configure(bg=BG_LIGHT, padx=20, pady=16)
        _lbl(p,"Question Bank",bold=True).grid(row=0,column=0,sticky="w",pady=(0,4))
        frm=tk.Frame(p,bg=BG_LIGHT); frm.grid(row=1,column=0,sticky="ew",pady=4)
        self._bp=tk.StringVar(value="No file selected")
        tk.Label(frm,textvariable=self._bp,bg=BG_LIGHT,fg=BG_DARK,
                 font=FB,width=60,anchor="w").pack(side="left",padx=(0,8))
        _btn(frm,"Browse…",self._browse_bank).pack(side="left")
        self._bs=tk.StringVar(value="")
        tk.Label(p,textvariable=self._bs,bg=BG_LIGHT,fg=BG_MID,
                 font=FB,justify="left").grid(row=2,column=0,sticky="w",pady=6)
        _lbl(p,"(الناتج, المؤشر) Pair Summary",bold=True).grid(
            row=3,column=0,sticky="w",pady=(12,4))
        frm2=tk.Frame(p,bg=BG_LIGHT); frm2.grid(row=4,column=0,sticky="nsew")
        self._ptree=self._treeview(frm2,
            ["الناتج","المؤشر","Count","Mean b","Min b","Max b"],
            [180,180,70,80,80,80])

    def _browse_bank(self):
        path=fd.askopenfilename(title="Select Bank",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not path: return
        self._status("Loading…")
        threading.Thread(target=self._load_bank_t, args=(path,), daemon=True).start()

    def _load_bank_t(self, path):
        try:
            df=load_bank(path); self._q.put(("bank_ok",df,path))
        except Exception as e:
            self._q.put(("err",str(e)))

    def _on_bank(self, df, path):
        self.bank_df=df
        self._bp.set(Path(path).name)
        self._bs.set(
            f"Loaded {len(df)} questions  |  Pairs: {df['_pair_key'].nunique()}  |  "
            f"b range: [{df['b'].min():.2f},{df['b'].max():.2f}]  |  Mean b: {df['b'].mean():.3f}")
        summ=get_pair_summary(df)
        for i in self._ptree.get_children(): self._ptree.delete(i)
        for _,row in summ.iterrows():
            self._ptree.insert("","end",values=[
                row["الناتج"],row["المؤشر"],row["count"],
                f"{row['mean_b']:.3f}",f"{row['min_b']:.3f}",f"{row['max_b']:.3f}"])
        self._pop_pairs(df)
        self._status(f"Bank loaded: {len(df)} questions.")

    # ══════════════ TAB 2 — CONFIGURE ══════════════
    def _build_config(self, p):
        p.configure(bg=BG_LIGHT, padx=20, pady=16)
        p.columnconfigure(0,weight=1); p.columnconfigure(1,weight=1)
        _lbl(p,"Assembly Settings",bold=True).grid(row=0,column=0,columnspan=2,sticky="w",pady=(0,6))

        # basic params
        pf=tk.LabelFrame(p,text="Basic Parameters",bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=10,pady=8)
        pf.grid(row=1,column=0,sticky="nsew",padx=(0,8),pady=4)
        self._nf  =self._le(pf,"Number of Forms:",0,"4")
        self._qpf =self._le(pf,"Questions / Form:",1,"40")
        self._mb  =self._le(pf,"Target Mean b:",2,"0.0")
        self._minb=self._le(pf,"Min b:",3,"-3.0")
        self._maxb=self._le(pf,"Max b:",4,"3.0")
        self._seed=self._le(pf,"RNG Seed (blank=random):",5,"")
        self._reuse=tk.BooleanVar(value=False)
        tk.Checkbutton(pf,text="Allow reuse across forms",variable=self._reuse,
                       bg=BG_LIGHT,fg=TEXT_DARK,font=FB).grid(row=6,column=0,columnspan=2,sticky="w",pady=4)

        # weights
        wf=tk.LabelFrame(p,text="Scoring Weights (sum≈1.0)",bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=10,pady=8)
        wf.grid(row=1,column=1,sticky="nsew",padx=(8,0),pady=4)
        self._wd=self._le(wf,"Difficulty Fit:",0,"0.40")
        self._wb=self._le(wf,"Bin Fit:",1,"0.25")
        self._wu=self._le(wf,"Usage Penalty:",2,"0.20")
        self._wr=self._le(wf,"Randomness:",3,"0.15")

        # simulation
        sf=tk.LabelFrame(p,text="Simulation",bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=10,pady=8)
        sf.grid(row=2,column=0,sticky="nsew",padx=(0,8),pady=4)
        self._sn=self._le(sf,"Students:",0,"3000")
        self._tm=self._le(sf,"θ mean:",1,"0.0")
        self._ts=self._le(sf,"θ SD:",2,"1.0")

        # output
        of=tk.LabelFrame(p,text="Output",bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=10,pady=8)
        of.grid(row=2,column=1,sticky="nsew",padx=(8,0),pady=4)
        self._odir=tk.StringVar(value=str(Path.home()/"TestOutput"))
        tk.Entry(of,textvariable=self._odir,width=30,font=FB,bg="#FFF",fg="#000").grid(row=0,column=0,sticky="ew",pady=4)
        _btn(of,"Browse…",self._browse_out).grid(row=0,column=1,padx=(6,0))
        of.columnconfigure(0,weight=1)
        self._pfx=self._le(of,"Prefix:",1,"test")

        # bins
        bf=tk.LabelFrame(p,text="Difficulty Bins (proportions must sum to 1.0)",
                          bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=10,pady=8)
        bf.grid(row=3,column=0,columnspan=2,sticky="ew",pady=6)
        self._build_bins(bf)

        # pair requirements
        pr=tk.LabelFrame(p,text="(الناتج , المؤشر) Pair Requirements",
                          bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=10,pady=8)
        pr.grid(row=4,column=0,columnspan=2,sticky="ew",pady=6)
        self._build_pair_sec(pr)

    def _le(self, parent, label, row, default=""):
        tk.Label(parent,text=label,bg=parent.cget("bg"),fg=TEXT_DARK,
                 font=FB,anchor="w").grid(row=row,column=0,sticky="w",padx=(0,8),pady=2)
        e=_e(parent,w=14); e.insert(0,default)
        e.grid(row=row,column=1,sticky="w",pady=2)
        return e

    def _build_bins(self, parent):
        hdr=tk.Frame(parent,bg=BG_LIGHT); hdr.pack(fill="x",pady=(0,4))
        for i,(t,w) in enumerate([("Low b",8),("High b",8),("Proportion",10)]):
            tk.Label(hdr,text=t,bg=BG_MID,fg=TEXT_LT,font=FH,width=w,
                     relief="flat",padx=4).grid(row=0,column=i,padx=2)
        self._bframe=tk.Frame(parent,bg=BG_LIGHT); self._bframe.pack(fill="x")
        self._brows: list = []
        ctrl=tk.Frame(parent,bg=BG_LIGHT); ctrl.pack(fill="x",pady=(6,0))
        _btn(ctrl,"+ Add Bin",self._add_bin,bg=BG_MID).pack(side="left",padx=4)
        _btn(ctrl,"Clear",self._clr_bins,bg="#C00000").pack(side="left",padx=4)
        _btn(ctrl,"Auto 3",lambda:self._auto_bins(3),bg=BG_DARK).pack(side="left",padx=4)
        for lo,hi,pr in [(-3.0,-0.5,0.25),(-0.5,0.5,0.50),(0.5,3.0,0.25)]:
            self._add_bin(lo,hi,pr)

    def _add_bin(self, lo=-1.0, hi=1.0, pr=0.33):
        frm=tk.Frame(self._bframe,bg=BG_LIGHT); frm.pack(fill="x",pady=2)
        elo=_e(frm,w=10); elo.insert(0,str(lo))
        ehi=_e(frm,w=10); ehi.insert(0,str(hi))
        epr=_e(frm,w=10); epr.insert(0,str(pr))
        elo.grid(row=0,column=0,padx=4); ehi.grid(row=0,column=1,padx=4)
        epr.grid(row=0,column=2,padx=4)
        t=(elo,ehi,epr); self._brows.append(t)
        dx=_btn(frm,"×",lambda rt=t,f=frm:self._del_bin(rt,f),bg="#C00000")
        dx.config(width=2,pady=2); dx.grid(row=0,column=3,padx=4)

    def _del_bin(self, t, frm):
        if t in self._brows: self._brows.remove(t)
        frm.destroy()

    def _clr_bins(self):
        for w in self._bframe.winfo_children(): w.destroy()
        self._brows.clear()

    def _auto_bins(self, n):
        self._clr_bins()
        lo=float(self._minb.get() or -3.0); hi=float(self._maxb.get() or 3.0)
        step=(hi-lo)/n; pr=round(1/n,4)
        for i in range(n):
            self._add_bin(round(lo+i*step,3), round(lo+(i+1)*step,3), pr)

    def _build_pair_sec(self, parent):
        tk.Label(parent,text="Load a bank first to auto-populate.",
                 bg=BG_LIGHT,fg=BG_MID,font=FB).pack(anchor="w",pady=(0,4))
        hdr=tk.Frame(parent,bg=BG_LIGHT); hdr.pack(fill="x",pady=(0,2))
        for i,(t,w) in enumerate([("الناتج",20),("المؤشر",20),("Count",8),("Available",10)]):
            tk.Label(hdr,text=t,bg=BG_MID,fg=TEXT_LT,font=FH,width=w,
                     relief="flat",padx=4).grid(row=0,column=i,padx=2)
        self._pframe=tk.Frame(parent,bg=BG_LIGHT); self._pframe.pack(fill="x")
        self._prows: list[dict] = []
        self._ptot=tk.Label(parent,text="",bg=BG_LIGHT,fg=BG_DARK,font=FB)
        self._ptot.pack(anchor="w",pady=4)

    def _pop_pairs(self, df):
        for w in self._pframe.winfo_children(): w.destroy()
        self._prows.clear()
        summ=get_pair_summary(df)
        n=len(summ)
        if n==0: return
        try: qpf=int(self._qpf.get())
        except: qpf=40
        base=qpf//n; rem=qpf%n
        counts=[base+1 if i<rem else base for i in range(n)]
        for idx,(_,row) in enumerate(summ.iterrows()):
            frm=tk.Frame(self._pframe,bg=BG_LIGHT); frm.pack(fill="x",pady=1)
            en=tk.Entry(frm,width=22,bg="#F0F0F0",fg=TEXT_DARK,font=FB,state="readonly")
            em=tk.Entry(frm,width=22,bg="#F0F0F0",fg=TEXT_DARK,font=FB,state="readonly")
            ec=_e(frm,w=8)
            ea=tk.Entry(frm,width=10,bg="#F0F0F0",fg=TEXT_DARK,font=FB,state="readonly")
            en.config(state="normal"); en.insert(0,str(row["الناتج"])); en.config(state="readonly")
            em.config(state="normal"); em.insert(0,str(row["المؤشر"])); em.config(state="readonly")
            ec.insert(0,str(counts[idx]))
            ea.config(state="normal"); ea.insert(0,str(int(row["count"]))); ea.config(state="readonly")
            en.grid(row=0,column=0,padx=4); em.grid(row=0,column=1,padx=4)
            ec.grid(row=0,column=2,padx=4); ea.grid(row=0,column=3,padx=4)
            self._prows.append({
                "pair_key": str(row["الناتج"])+"||"+str(row["المؤشر"]),
                "en":en,"em":em,"ec":ec,"ea":ea})
        self._upd_pair_total()

    def _upd_pair_total(self):
        total=sum(int(r["ec"].get()) for r in self._prows
                  if r["ec"].get().strip().isdigit())
        try: qpf=int(self._qpf.get())
        except: qpf=40
        col="green" if total==qpf else "#C00000"
        self._ptot.config(text=f"Total: {total}  (target: {qpf})",fg=col)

    def _browse_out(self):
        d=fd.askdirectory(title="Output Directory")
        if d: self._odir.set(d)

    # ══════════════ TAB 3 — GENERATE ══════════════
    def _build_generate(self, p):
        p.configure(bg=BG_LIGHT)
        tk.Frame(p,bg=BG_LIGHT,pady=14).pack(fill="x",padx=20)
        _lbl(p,"Generate Forms & Run Psychometric Analysis",bold=True).pack(anchor="w",padx=20)
        bf=tk.Frame(p,bg=BG_LIGHT,pady=6); bf.pack(padx=20,anchor="w")
        self._gbtn=_btn(bf,"▶  Generate & Analyse",self._run,bg=ACCENT)
        self._gbtn.config(font=("Segoe UI",11,"bold"),pady=8,padx=20)
        self._gbtn.pack(side="left",padx=(0,12))
        self._pv=tk.DoubleVar(value=0)
        ttk.Progressbar(bf,variable=self._pv,maximum=100,
                        length=300,mode="determinate").pack(side="left")
        lf=tk.LabelFrame(p,text="Log",bg=BG_LIGHT,fg=BG_DARK,font=FH,padx=8,pady=8)
        lf.pack(fill="both",expand=True,padx=20,pady=8)
        self._log=tk.Text(lf,height=20,bg="#0D1117",fg="#58D68D",
                          font=FM,state="disabled",relief="flat")
        sb=ttk.Scrollbar(lf,command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); self._log.pack(fill="both",expand=True)

    def _run(self):
        if self.bank_df is None:
            mb.showerror("No Bank","Load a bank first (Tab 1)."); return
        try: params=self._params()
        except ValueError as e:
            mb.showerror("Config Error",str(e)); return
        self._gbtn.config(state="disabled"); self._pv.set(0)
        self._clr_log()
        threading.Thread(target=self._gen_t, args=(params,), daemon=True).start()

    def _gen_t(self, params):
        try:
            self._lg("Building config…")
            cfg=build_config_from_gui(params)
            self._lg(f"Assembling {cfg.n_forms} × {cfg.questions_per_form} questions…")
            eng=AssemblyEngine(self.bank_df, cfg)
            res=eng.assemble()
            self._q.put(("prog",40))
            for w in res.warnings: self._lg(f"  ⚠ {w}","yellow")
            self._lg("Psychometric analysis…")
            sim=SimulationConfig(n_students=params["n_students"],
                                 theta_mean=params["theta_mean"],
                                 theta_sd=params["theta_sd"],
                                 rng_seed=params["rng_seed"])
            names=[f"Form {i+1}" for i in range(len(res.forms))]
            analyses=analyse_all_forms(res.forms, names, sim_cfg=sim)
            self._q.put(("prog",75))
            self._lg("Writing output files…")
            paths=generate_all_outputs(analyses, Path(params["output_dir"]),
                                       prefix=params["prefix"])
            self._q.put(("prog",100))
            for lbl,p in paths.items(): self._lg(f"  [{lbl}]  {p}")
            self._lg("\n✓ Done!","lime")
            self._q.put(("done", analyses))
        except Exception as e:
            self._lg(f"\n✗ Error: {e}","red")
            logger.exception("Generation failed")
            self._q.put(("fail",))

    def _params(self):
        def fi(e, n):
            try: return float(e.get().strip())
            except: raise ValueError(f"Invalid '{n}'")
        def ii(e, n):
            try: return int(e.get().strip())
            except: raise ValueError(f"Invalid '{n}'")
        nf=ii(self._nf,"Forms"); qpf=ii(self._qpf,"Q/Form")
        mb_=fi(self._mb,"Mean b"); minb=fi(self._minb,"Min b"); maxb=fi(self._maxb,"Max b")
        sd=self._seed.get().strip(); seed=int(sd) if sd else None
        bins=[]
        for elo,ehi,epr in self._brows:
            try: bins.append({"low":float(elo.get()),"high":float(ehi.get()),
                               "proportion":float(epr.get())})
            except: raise ValueError("Invalid bin value")
        if bins:
            tp=sum(b["proportion"] for b in bins)
            if not 0.98<=tp<=1.02:
                raise ValueError(f"Bin proportions sum to {tp:.4f}, must be 1.0")
        pr=[]
        for r in self._prows:
            try: cnt=int(r["ec"].get())
            except: cnt=0
            pr.append({"pair_key":r["pair_key"],"الناتج":r["en"].get(),
                       "المؤشر":r["em"].get(),"count":cnt})
        if pr:
            tot=sum(r["count"] for r in pr)
            if tot!=qpf: raise ValueError(f"Pair total {tot} ≠ Q/Form {qpf}")
        return {"n_forms":nf,"questions_per_form":qpf,"target_mean_b":mb_,
                "target_min_b":minb,"target_max_b":maxb,"bins":bins,
                "pair_requirements":pr,
                "w_difficulty":fi(self._wd,"w_diff"),"w_bin":fi(self._wb,"w_bin"),
                "w_usage":fi(self._wu,"w_usage"),"w_random":fi(self._wr,"w_rand"),
                "allow_reuse_across_forms":self._reuse.get(),"rng_seed":seed,
                "output_dir":self._odir.get(),
                "prefix":self._pfx.get() or "test",
                "n_students":ii(self._sn,"students"),
                "theta_mean":fi(self._tm,"θ mean"),"theta_sd":fi(self._ts,"θ SD")}

    # ══════════════ TAB 4 — RESULTS ══════════════
    def _build_results(self, p):
        p.configure(bg=BG_LIGHT,padx=20,pady=16)
        p.columnconfigure(0,weight=1)
        _lbl(p,"Results Summary",bold=True).grid(row=0,column=0,sticky="w",pady=(0,8))
        self._rtree=self._treeview(p,
            ["Form","N","Mean b","SD b","Min b","Max b","Mean a","Mean c","Alpha","Corr."],
            [90,50,80,70,70,70,70,70,70,80],row=1)
        bf=tk.Frame(p,bg=BG_LIGHT); bf.grid(row=2,column=0,sticky="w",pady=8)
        _btn(bf,"Open Folder",self._open_folder).pack(side="left")
        _btn(bf,"Refresh",self._refresh,bg=BG_DARK).pack(side="left",padx=6)

    def _refresh(self):
        if not self.analyses: return
        for i in self._rtree.get_children(): self._rtree.delete(i)
        for fa in self.analyses:
            s=fa.summary
            self._rtree.insert("","end",values=[
                fa.form_name,s.get("n_items",""),
                f"{s.get('mean_b',0):.3f}",f"{s.get('std_b',0):.3f}",
                f"{s.get('min_b',0):.3f}",f"{s.get('max_b',0):.3f}",
                f"{s.get('mean_a',0):.3f}",f"{s.get('mean_c',0):.3f}",
                f"{s.get('alpha',0):.3f}",f"{s.get('mean_item_cor',0):.3f}"])

    def _open_folder(self):
        d=self._odir.get()
        if sys.platform=="win32": subprocess.Popen(["explorer",d])
        elif sys.platform=="darwin": subprocess.Popen(["open",d])
        else: subprocess.Popen(["xdg-open",d])

    # ══════════════ TREEVIEW ══════════════
    def _treeview(self, parent, columns, widths, row=0):
        s=ttk.Style()
        s.configure("C.Treeview.Heading",background=BG_MID,foreground=TEXT_LT,font=FH)
        s.configure("C.Treeview",font=FB,rowheight=22)
        s.map("C.Treeview",background=[("selected",ACCENT)])
        frm=tk.Frame(parent,bg=BG_LIGHT)
        frm.grid(row=row,column=0,sticky="nsew")
        tree=ttk.Treeview(frm,columns=columns,show="headings",style="C.Treeview")
        vsb=ttk.Scrollbar(frm,orient="vertical",command=tree.yview)
        hsb=ttk.Scrollbar(frm,orient="horizontal",command=tree.xview)
        tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        for col,w in zip(columns,widths):
            tree.heading(col,text=col); tree.column(col,width=w,minwidth=40,anchor="center")
        vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tree.grid(row=0,column=0,sticky="nsew")
        frm.columnconfigure(0,weight=1); frm.rowconfigure(0,weight=1)
        return tree

    # ══════════════ LOG HELPERS ══════════════
    def _lg(self, msg, col=""):
        self._q.put(("log",msg,col))

    def _clr_log(self):
        self._q.put(("logclr",))

    def _append_log(self, msg, col=""):
        self._log.config(state="normal")
        tag=col or "g"; self._log.tag_config(tag,foreground=col or "#58D68D")
        self._log.insert("end",msg+"\n",tag); self._log.see("end")
        self._log.config(state="disabled")

    # ══════════════ QUEUE POLL ══════════════
    def _poll(self):
        try:
            while True:
                item=self._q.get_nowait(); k=item[0]
                if k=="bank_ok": self._on_bank(item[1],item[2])
                elif k=="log":
                    self._append_log(item[1], item[2] if len(item)>2 else "")
                elif k=="logclr":
                    self._log.config(state="normal")
                    self._log.delete("1.0","end")
                    self._log.config(state="disabled")
                elif k=="prog": self._pv.set(item[1])
                elif k=="done":
                    self.analyses=item[1]
                    self._gbtn.config(state="normal"); self._pv.set(100)
                    self._refresh(); self.nb.select(3)
                    self._status("Done.")
                    mb.showinfo("Done","All forms generated and analysed.\nOutput files written.")
                elif k=="fail":
                    self._gbtn.config(state="normal")
                    self._status("Failed — see log.")
                elif k=="err":
                    self._status(f"Error: {item[1]}")
                    mb.showerror("Error",item[1])
        except queue.Empty:
            pass
        self.after(80, self._poll)


# ════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════

def _cli(args):
    bank=load_bank(args.bank)
    cfg=AssemblyConfig(n_forms=args.forms,questions_per_form=args.qpf,
                       target_mean_b=args.mean_b,target_min_b=args.min_b,
                       target_max_b=args.max_b,allow_reuse_across_forms=args.reuse,
                       rng_seed=args.seed)
    result=AssemblyEngine(bank,cfg).assemble()
    for w in result.warnings: logger.warning(w)
    sim=SimulationConfig(n_students=args.n_students,rng_seed=args.seed)
    names=[f"Form {i+1}" for i in range(len(result.forms))]
    analyses=analyse_all_forms(result.forms,names,sim_cfg=sim)
    paths=generate_all_outputs(analyses,Path(args.output),prefix=args.prefix)
    for lbl,p in paths.items(): logger.info("  %-15s %s",lbl,p)


if __name__ == "__main__":
    parser=argparse.ArgumentParser(description="Smart Test Assembly")
    parser.add_argument("--cli",action="store_true")
    parser.add_argument("--bank",type=str)
    parser.add_argument("--forms",type=int,default=4)
    parser.add_argument("--qpf",type=int,default=40)
    parser.add_argument("--mean-b",type=float,default=0.0,dest="mean_b")
    parser.add_argument("--min-b",type=float,default=-3.0,dest="min_b")
    parser.add_argument("--max-b",type=float,default=3.0,dest="max_b")
    parser.add_argument("--reuse",action="store_true")
    parser.add_argument("--seed",type=int,default=None)
    parser.add_argument("--n-students",type=int,default=3000,dest="n_students")
    parser.add_argument("--output",type=str,default="./output")
    parser.add_argument("--prefix",type=str,default="test")
    args=parser.parse_args()

    if args.cli:
        if not args.bank: parser.error("--bank required in CLI mode")
        _cli(args)
    else:
        app=App(); app.mainloop()
