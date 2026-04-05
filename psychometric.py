"""
Psychometric Analysis Module
-----------------------------
Full 3PL IRT model implementation with vectorised NumPy operations.

Provides:
    - 3PL probability function
    - Simulated response generation (3000+ examinees)
    - Item Characteristic Curves (ICC)
    - Test Characteristic Curve (TCC)
    - Test Information Function (TIF)
    - Cronbach's Alpha (from simulated responses)
    - Corrected item-total correlations
    - Per-form and cross-form summary statistics
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 3PL Model
# ---------------------------------------------------------------------------

def p3pl(theta: np.ndarray, a: np.ndarray, b: np.ndarray, c: np.ndarray) -> np.ndarray:
    """
    Vectorised 3PL probability.

    Parameters
    ----------
    theta : shape (N,) – ability values
    a     : shape (K,) – discrimination
    b     : shape (K,) – difficulty
    c     : shape (K,) – guessing

    Returns
    -------
    P : shape (N, K) – P(correct | theta, a, b, c)
    """
    theta = np.asarray(theta, dtype=float).reshape(-1, 1)  # (N, 1)
    a = np.asarray(a, dtype=float).reshape(1, -1)           # (1, K)
    b = np.asarray(b, dtype=float).reshape(1, -1)
    c = np.asarray(c, dtype=float).reshape(1, -1)

    exp_term = np.exp(-a * (theta - b))          # (N, K)
    P = c + (1.0 - c) / (1.0 + exp_term)
    return P                                      # (N, K)


def item_information(theta: np.ndarray, a: np.ndarray, b: np.ndarray, c: np.ndarray) -> np.ndarray:
    """
    Item information function under 3PL.
    Returns shape (N, K).
    """
    P = p3pl(theta, a, b, c)
    Q = 1.0 - P
    a_2d = np.asarray(a, dtype=float).reshape(1, -1)
    c_2d = np.asarray(c, dtype=float).reshape(1, -1)

    # IIF = a² * ((P - c) / (1 - c))² * (Q / P)
    numerator = ((P - c_2d) / (1.0 - c_2d + 1e-12)) ** 2
    info = (a_2d ** 2) * numerator * (Q / (P + 1e-12))
    return info                                   # (N, K)


# ---------------------------------------------------------------------------
# Simulation
# ---------------------------------------------------------------------------

@dataclass
class SimulationConfig:
    n_students: int = 3000
    theta_mean: float = 0.0
    theta_sd: float = 1.0
    theta_distribution: str = "normal"   # "normal" | "uniform"
    rng_seed: Optional[int] = None


def simulate_responses(
    form_df: pd.DataFrame,
    sim_cfg: Optional[SimulationConfig] = None,
) -> tuple[np.ndarray, np.ndarray]:
    """
    Simulate binary response matrix for a form.

    Returns
    -------
    theta   : (N,) examinee abilities
    X       : (N, K) binary response matrix  (1 = correct, 0 = incorrect)
    """
    if sim_cfg is None:
        sim_cfg = SimulationConfig()

    rng = np.random.default_rng(sim_cfg.rng_seed)
    N = sim_cfg.n_students
    K = len(form_df)

    if sim_cfg.theta_distribution == "uniform":
        theta = rng.uniform(-3.0, 3.0, size=N)
    else:
        theta = rng.normal(sim_cfg.theta_mean, sim_cfg.theta_sd, size=N)

    a = form_df["a"].to_numpy(dtype=float)
    b = form_df["b"].to_numpy(dtype=float)
    c = form_df["c"].to_numpy(dtype=float)

    P = p3pl(theta, a, b, c)            # (N, K)
    u = rng.random((N, K))
    X = (u < P).astype(np.int8)
    return theta, X


# ---------------------------------------------------------------------------
# Psychometric statistics
# ---------------------------------------------------------------------------

def cronbach_alpha(X: np.ndarray) -> float:
    """
    Cronbach's alpha from a binary response matrix (N × K).
    Uses the KR-20 formula (equivalent for binary items).
    """
    N, K = X.shape
    if K < 2:
        return float("nan")
    item_vars = X.var(axis=0, ddof=1)
    total_var = X.sum(axis=1).var(ddof=1)
    if total_var == 0:
        return float("nan")
    alpha = (K / (K - 1)) * (1.0 - item_vars.sum() / total_var)
    return float(np.clip(alpha, -1.0, 1.0))


def item_total_correlations(X: np.ndarray) -> np.ndarray:
    """
    Corrected item-total correlation for each item.
    Returns shape (K,).
    """
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


# ---------------------------------------------------------------------------
# Curve data generation
# ---------------------------------------------------------------------------

@dataclass
class FormAnalysis:
    form_name: str
    form_df: pd.DataFrame
    theta_grid: np.ndarray        # (T,) theta values for curves
    P_grid: np.ndarray            # (T, K) ICC values
    TCC: np.ndarray               # (T,) Test Characteristic Curve
    TIF: np.ndarray               # (T,) Test Information Function
    SEM: np.ndarray               # (T,) Standard Error of Measurement
    alpha: float
    item_cors: np.ndarray         # (K,)
    sim_theta: np.ndarray         # (N,)
    sim_X: np.ndarray             # (N, K)
    summary: dict


def analyse_form(
    form_df: pd.DataFrame,
    form_name: str = "Form",
    n_theta_points: int = 200,
    sim_cfg: Optional[SimulationConfig] = None,
) -> FormAnalysis:
    """
    Full psychometric analysis for a single form.
    """
    if form_df.empty:
        raise ValueError(f"Form '{form_name}' is empty.")

    a = form_df["a"].to_numpy(dtype=float)
    b = form_df["b"].to_numpy(dtype=float)
    c = form_df["c"].to_numpy(dtype=float)
    K = len(form_df)

    # Theta grid for curves
    theta_grid = np.linspace(-4.0, 4.0, n_theta_points)

    # ICC
    P_grid = p3pl(theta_grid, a, b, c)   # (T, K)

    # TCC = sum of ICC values across items
    TCC = P_grid.sum(axis=1)             # (T,)

    # TIF = sum of item information functions
    info_grid = item_information(theta_grid, a, b, c)   # (T, K)
    TIF = info_grid.sum(axis=1)                          # (T,)

    # SEM = 1 / sqrt(TIF)
    SEM = 1.0 / np.sqrt(np.clip(TIF, 1e-8, None))

    # Simulate responses
    theta_sim, X_sim = simulate_responses(form_df, sim_cfg)

    # Cronbach alpha
    alpha = cronbach_alpha(X_sim)

    # Item-total correlations
    item_cors = item_total_correlations(X_sim)

    # Summary statistics
    summary = {
        "n_items": K,
        "mean_b": float(np.mean(b)),
        "std_b": float(np.std(b)),
        "min_b": float(np.min(b)),
        "max_b": float(np.max(b)),
        "mean_a": float(np.mean(a)),
        "mean_c": float(np.mean(c)),
        "alpha": alpha,
        "mean_item_cor": float(np.nanmean(item_cors)),
        "min_item_cor": float(np.nanmin(item_cors)),
        "max_item_cor": float(np.nanmax(item_cors)),
        "peak_info_theta": float(theta_grid[np.argmax(TIF)]),
        "peak_info": float(np.max(TIF)),
    }

    return FormAnalysis(
        form_name=form_name,
        form_df=form_df,
        theta_grid=theta_grid,
        P_grid=P_grid,
        TCC=TCC,
        TIF=TIF,
        SEM=SEM,
        alpha=alpha,
        item_cors=item_cors,
        sim_theta=theta_sim,
        sim_X=X_sim,
        summary=summary,
    )


def analyse_all_forms(
    forms: list[pd.DataFrame],
    form_names: Optional[list[str]] = None,
    sim_cfg: Optional[SimulationConfig] = None,
) -> list[FormAnalysis]:
    """Run analyse_form for every form."""
    if form_names is None:
        form_names = [f"Form {i + 1}" for i in range(len(forms))]

    analyses: list[FormAnalysis] = []
    for df, name in zip(forms, form_names):
        try:
            fa = analyse_form(df, form_name=name, sim_cfg=sim_cfg)
            analyses.append(fa)
        except Exception as exc:
            logger.error("Analysis failed for '%s': %s", name, exc)
    return analyses


def cross_form_summary(analyses: list[FormAnalysis]) -> pd.DataFrame:
    """Create a cross-form comparison DataFrame."""
    rows = []
    for fa in analyses:
        row = {"Form": fa.form_name}
        row.update(fa.summary)
        rows.append(row)
    return pd.DataFrame(rows)
