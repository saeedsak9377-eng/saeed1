"""
Assembly Engine
---------------
Intelligent, scoring-based test assembly.

Key design decisions
====================
* No random sampling as the primary strategy – questions are scored and
  selected deterministically (greedy), with a small stochastic noise term
  to prevent identical forms across runs while still favouring the best items.
* The (الناتج, المؤشر) pair is treated as an independent entity throughout.
* "Bins" are contiguous difficulty sub-ranges; each bin has an exact integer
  quota computed via the Largest Remainder Method to avoid rounding drift.
* Usage tracking persists across forms within a session so that no question
  is over-represented.
* Shortage handling uses adaptive constraint relaxation instead of placeholders:
    Stage 1 – strict (b within bin range, exact pair match)
    Stage 2 – relax mean tolerance by 50 %
    Stage 3 – relax bins by ±0.2 and allow pair-level substitution
    Stage 4 – global pool fallback (least-used, closest difficulty)

Used-Blocks fallback (last resort)
====================================
When the main unused question pool for a target block is fully exhausted,
_fallback_select may reuse already-used questions from the "Used Blocks"
dataset (i.e. items whose usage_count > 0).

Strict Stage-Level matching rule
---------------------------------
Reuse is restricted exclusively to questions that share the same Stage AND
Level/Part as the block currently being assembled.  Matching is based on the
block's QuestionID prefix (the leading dot-delimited segments that encode
stage and level, e.g. "GAT-2.1" from "GAT-2.1.xxxxx").

The number of prefix segments used for matching is controlled by
AssemblyConfig.block_prefix_segments (default 2).  Given a QuestionID like
"GAT-2.1.045", splitting on "." yields ["GAT-2", "1", "045"]; the first 2
segments joined back with "." give the prefix "GAT-2.1", which is what
every candidate in the fallback pool must also start with.

Cross-stage or cross-level borrowing is never permitted under any
circumstance, even when the same-prefix reuse pool is also exhausted
(in that case the fallback simply returns nothing rather than violating
the rule).
"""

from __future__ import annotations

import logging
import math
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class BinSpec:
    low: float
    high: float
    count: int          # exact integer quota (post-LRM)


@dataclass
class PairSpec:
    pair_key: str
    الناتج: str
    المؤشر: str
    count: int          # total questions needed from this pair


@dataclass
class AssemblyConfig:
    n_forms: int = 1
    questions_per_form: int = 40

    # Difficulty target
    target_mean_b: float = 0.0
    target_min_b: float = -3.0
    target_max_b: float = 3.0

    # Bins (list of (low, high, proportion) tuples)
    bins: list[tuple[float, float, float]] = field(default_factory=list)

    # (الناتج, المؤشر) pair requirements list[dict]
    # Each dict: {"pair_key": str, "الناتج": str, "المؤشر": str, "count": int}
    pair_requirements: list[dict] = field(default_factory=list)

    # Scoring weights
    w_difficulty: float = 0.40
    w_bin: float        = 0.25
    w_usage: float      = 0.20
    w_random: float     = 0.15

    # Assembly options
    allow_reuse_across_forms: bool = False
    relaxation_stages: int = 4
    rng_seed: Optional[int] = None

    # Used-Blocks fallback: number of leading dot-delimited QuestionID
    # segments that define a block's Stage-Level prefix for strict matching.
    # E.g. with block_prefix_segments=2, "GAT-2.1.045" → prefix "GAT-2.1".
    block_prefix_segments: int = 2


@dataclass
class AssemblyResult:
    forms: list[pd.DataFrame]          # one DataFrame per form
    stats: list[dict]                  # per-form statistics
    warnings: list[str]
    usage_counts: dict[str, int]       # QuestionID → times used


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _largest_remainder(total: int, proportions: list[float]) -> list[int]:
    """Allocate *total* integers to buckets by proportion using LRM."""
    if not proportions:
        return []
    floats = [total * p / sum(proportions) for p in proportions]
    floors = [int(f) for f in floats]
    remainders = [(f - fl, i) for i, (f, fl) in enumerate(zip(floats, floors))]
    deficit = total - sum(floors)
    remainders.sort(reverse=True)
    for k in range(deficit):
        floors[remainders[k][1]] += 1
    return floors


def _irt_b_score(b_values: np.ndarray, target_mean: float, b_range: tuple[float, float]) -> np.ndarray:
    """
    Score items by how close their difficulty is to the target mean.
    Returns values in [0, 1], higher = better fit.
    """
    span = max(b_range[1] - b_range[0], 1e-6)
    dist = np.abs(b_values - target_mean)
    return 1.0 - np.clip(dist / span, 0.0, 1.0)


def _bin_membership_score(b_values: np.ndarray, bin_low: float, bin_high: float) -> np.ndarray:
    """1 if b is within bin, else decaying score based on distance from bin."""
    mid = (bin_low + bin_high) / 2.0
    half_width = max((bin_high - bin_low) / 2.0, 1e-6)
    dist = np.abs(b_values - mid)
    return np.clip(1.0 - (dist - half_width) / half_width, 0.0, 1.0)


class AssemblyEngine:
    """
    Assembles multiple test forms from a question bank using a
    scoring-based greedy algorithm.
    """

    def __init__(self, bank: pd.DataFrame, config: AssemblyConfig):
        self.bank = bank.copy().reset_index(drop=True)
        self.config = config
        self.rng = np.random.default_rng(config.rng_seed)

        # usage_counts[QuestionID] = number of forms it has appeared in
        self.usage_counts: dict[str, int] = {
            qid: 0 for qid in self.bank["QuestionID"].astype(str)
        }
        self.warnings: list[str] = []

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def assemble(self) -> AssemblyResult:
        cfg = self.config
        forms: list[pd.DataFrame] = []
        stats: list[dict] = []

        resolved_bins = self._resolve_bins()
        resolved_pairs = self._resolve_pairs()

        for form_idx in range(cfg.n_forms):
            logger.info("Assembling form %d / %d", form_idx + 1, cfg.n_forms)
            form_df, form_stats = self._assemble_one_form(
                form_idx, resolved_bins, resolved_pairs
            )
            forms.append(form_df)
            stats.append(form_stats)

        return AssemblyResult(
            forms=forms,
            stats=stats,
            warnings=self.warnings,
            usage_counts=dict(self.usage_counts),
        )

    # ------------------------------------------------------------------
    # Bin & pair resolution
    # ------------------------------------------------------------------

    def _resolve_bins(self) -> list[BinSpec]:
        cfg = self.config
        n = cfg.questions_per_form

        if not cfg.bins:
            # Single implicit bin spanning the full difficulty range
            return [BinSpec(low=cfg.target_min_b, high=cfg.target_max_b, count=n)]

        proportions = [prop for _, _, prop in cfg.bins]
        counts = _largest_remainder(n, proportions)

        bins: list[BinSpec] = []
        for (low, high, _), cnt in zip(cfg.bins, counts):
            bins.append(BinSpec(low=low, high=high, count=cnt))

        allocated = sum(b.count for b in bins)
        if allocated != n:
            # Should not happen with LRM, but guard anyway
            bins[-1].count += n - allocated

        return bins

    def _resolve_pairs(self) -> list[PairSpec]:
        cfg = self.config
        n = cfg.questions_per_form

        if not cfg.pair_requirements:
            return []

        total_requested = sum(r["count"] for r in cfg.pair_requirements)
        if total_requested != n:
            # Scale proportionally
            proportions = [r["count"] for r in cfg.pair_requirements]
            counts = _largest_remainder(n, proportions)
            adjusted: list[PairSpec] = []
            for req, cnt in zip(cfg.pair_requirements, counts):
                adjusted.append(
                    PairSpec(
                        pair_key=req["pair_key"],
                        الناتج=req.get("الناتج", ""),
                        المؤشر=req.get("المؤشر", ""),
                        count=cnt,
                    )
                )
            return adjusted

        return [
            PairSpec(
                pair_key=r["pair_key"],
                الناتج=r.get("الناتج", ""),
                المؤشر=r.get("المؤشر", ""),
                count=r["count"],
            )
            for r in cfg.pair_requirements
        ]

    # ------------------------------------------------------------------
    # Single-form assembly
    # ------------------------------------------------------------------

    def _assemble_one_form(
        self,
        form_idx: int,
        bins: list[BinSpec],
        pairs: list[PairSpec],
    ) -> tuple[pd.DataFrame, dict]:
        cfg = self.config
        selected_ids: set[str] = set()
        selected_rows: list[pd.DataFrame] = []

        # ---------------------------------------------------------------
        # Strategy: iterate over pairs, for each pair iterate over bins
        # If no pair requirements, iterate over bins only
        # ---------------------------------------------------------------
        if pairs:
            selected_rows, selected_ids = self._assemble_by_pairs_and_bins(
                pairs, bins, form_idx
            )
        else:
            selected_rows, selected_ids = self._assemble_by_bins_only(
                bins, form_idx
            )

        form_df = pd.concat(selected_rows, ignore_index=True) if selected_rows else pd.DataFrame()

        # Update usage
        for qid in selected_ids:
            self.usage_counts[str(qid)] = self.usage_counts.get(str(qid), 0) + 1

        form_stats = self._compute_form_stats(form_df)
        return form_df, form_stats

    # ------------------------------------------------------------------
    # Pair × Bin assembly
    # ------------------------------------------------------------------

    def _assemble_by_pairs_and_bins(
        self,
        pairs: list[PairSpec],
        bins: list[BinSpec],
        form_idx: int,
    ) -> tuple[list[pd.DataFrame], set[str]]:
        cfg = self.config
        selected_ids: set[str] = set()
        all_rows: list[pd.DataFrame] = []

        # Distribute bin quotas across pairs proportionally
        pair_bin_matrix = self._distribute_pairs_to_bins(pairs, bins)

        for pair_idx, pair in enumerate(pairs):
            pair_pool = self._get_pool(pair.pair_key, selected_ids)

            # Derive the Stage-Level block prefix from this pair's question IDs.
            # This prefix is passed to _fallback_select so that, if last-resort
            # reuse from the Used-Blocks dataset is needed, only questions from
            # the exact same Stage and Level/Part are eligible.
            block_prefix = self._block_prefix_for_pair(pair.pair_key)

            for bin_idx, bin_spec in enumerate(bins):
                needed = pair_bin_matrix[pair_idx][bin_idx]
                if needed == 0:
                    continue

                rows, ids = self._select_from_pool(
                    pool=pair_pool,
                    needed=needed,
                    bin_spec=bin_spec,
                    already_selected=selected_ids,
                    form_idx=form_idx,
                )
                all_rows.extend(rows)
                selected_ids.update(ids)
                # Refresh pool to exclude just-selected
                pair_pool = pair_pool[~pair_pool["QuestionID"].isin(selected_ids)]

                if len(ids) < needed:
                    shortage = needed - len(ids)
                    self.warnings.append(
                        f"Form {form_idx + 1}, pair '{pair.pair_key}', "
                        f"bin [{bin_spec.low:.2f},{bin_spec.high:.2f}]: "
                        f"shortage of {shortage}. Applying fallback."
                    )
                    fallback_rows, fallback_ids = self._fallback_select(
                        shortage=shortage,
                        bin_spec=bin_spec,
                        already_selected=selected_ids,
                        form_idx=form_idx,
                        block_prefix=block_prefix,
                    )
                    all_rows.extend(fallback_rows)
                    selected_ids.update(fallback_ids)

        return all_rows, selected_ids

    def _assemble_by_bins_only(
        self,
        bins: list[BinSpec],
        form_idx: int,
    ) -> tuple[list[pd.DataFrame], set[str]]:
        selected_ids: set[str] = set()
        all_rows: list[pd.DataFrame] = []

        for bin_spec in bins:
            pool = self._get_pool(pair_key=None, exclude=selected_ids)
            rows, ids = self._select_from_pool(
                pool=pool,
                needed=bin_spec.count,
                bin_spec=bin_spec,
                already_selected=selected_ids,
                form_idx=form_idx,
            )
            all_rows.extend(rows)
            selected_ids.update(ids)

            if len(ids) < bin_spec.count:
                shortage = bin_spec.count - len(ids)
                self.warnings.append(
                    f"Form {form_idx + 1}, bin [{bin_spec.low:.2f},{bin_spec.high:.2f}]: "
                    f"shortage of {shortage}. Applying fallback."
                )
                fallback_rows, fallback_ids = self._fallback_select(
                    shortage=shortage,
                    bin_spec=bin_spec,
                    already_selected=selected_ids,
                    form_idx=form_idx,
                    block_prefix=None,  # no pair context; prefix matching not applicable
                )
                all_rows.extend(fallback_rows)
                selected_ids.update(fallback_ids)

        return all_rows, selected_ids

    # ------------------------------------------------------------------
    # Pool management
    # ------------------------------------------------------------------

    def _get_pool(
        self, pair_key: Optional[str], exclude: set[str]
    ) -> pd.DataFrame:
        """Return eligible items (not excluded, not over-used)."""
        cfg = self.config
        pool = self.bank.copy()

        # Filter by pair
        if pair_key is not None:
            pool = pool[pool["_pair_key"] == pair_key]

        # Exclude already-selected in this form
        if exclude:
            pool = pool[~pool["QuestionID"].isin(exclude)]

        # Penalise (but don't exclude) items used in previous forms
        # unless reuse is explicitly disallowed
        if not cfg.allow_reuse_across_forms:
            already_used = {
                qid for qid, cnt in self.usage_counts.items() if cnt > 0
            }
            pool = pool[~pool["QuestionID"].isin(already_used)]

        return pool.reset_index(drop=True)

    # ------------------------------------------------------------------
    # Core scoring & selection
    # ------------------------------------------------------------------

    def _score_items(
        self,
        pool: pd.DataFrame,
        bin_spec: BinSpec,
    ) -> np.ndarray:
        """
        Composite score for each item in pool:
            score = w_difficulty * difficulty_fit
                  + w_bin        * bin_fit
                  + w_usage      * (1 - normalised_usage)
                  + w_random     * noise
        """
        cfg = self.config
        b = pool["b"].to_numpy(dtype=float)
        qids = pool["QuestionID"].astype(str).tolist()

        # Difficulty fit: proximity to target mean
        diff_score = _irt_b_score(b, cfg.target_mean_b, (cfg.target_min_b, cfg.target_max_b))

        # Bin fit
        bin_score = _bin_membership_score(b, bin_spec.low, bin_spec.high)

        # Usage penalty (lower usage = higher score)
        usage_vals = np.array([self.usage_counts.get(qid, 0) for qid in qids], dtype=float)
        max_usage = max(usage_vals.max(), 1.0)
        usage_score = 1.0 - (usage_vals / max_usage)

        # Small random perturbation
        noise = self.rng.random(len(pool))

        scores = (
            cfg.w_difficulty * diff_score
            + cfg.w_bin      * bin_score
            + cfg.w_usage    * usage_score
            + cfg.w_random   * noise
        )
        return scores

    def _select_from_pool(
        self,
        pool: pd.DataFrame,
        needed: int,
        bin_spec: BinSpec,
        already_selected: set[str],
        form_idx: int,
    ) -> tuple[list[pd.DataFrame], set[str]]:
        """Select `needed` items from `pool` that fall within bin, greedy by score."""
        if pool.empty or needed == 0:
            return [], set()

        # Stage 1: strict bin membership
        candidates = pool[
            (pool["b"] >= bin_spec.low) & (pool["b"] <= bin_spec.high)
        ].copy()

        if len(candidates) < needed:
            # Stage 2: relax bin edges by 20%
            half_w = (bin_spec.high - bin_spec.low) / 2.0
            relaxed_low  = bin_spec.low  - 0.2 * half_w
            relaxed_high = bin_spec.high + 0.2 * half_w
            candidates = pool[
                (pool["b"] >= relaxed_low) & (pool["b"] <= relaxed_high)
            ].copy()

        if candidates.empty:
            # Stage 3: use closest-difficulty items from full pool
            candidates = pool.copy()

        scores = self._score_items(candidates, bin_spec)
        order = np.argsort(-scores)  # descending
        top = candidates.iloc[order[:needed]]

        selected_ids = set(top["QuestionID"].astype(str).tolist())
        return [top], selected_ids

    # ------------------------------------------------------------------
    # Block-prefix helpers (Stage-Level matching for Used-Blocks fallback)
    # ------------------------------------------------------------------

    @staticmethod
    def _derive_prefix(question_id: str, segments: int) -> str:
        """
        Return the leading *segments* dot-delimited parts of *question_id*.

        Examples (segments=2):
            "GAT-2.1.045"  → "GAT-2.1"
            "GAT-1.2.007"  → "GAT-1.2"
            "PLAIN"        → "PLAIN"   (fewer parts than requested → whole ID)
        """
        parts = str(question_id).split(".")
        return ".".join(parts[:segments])

    def _block_prefix_for_pair(self, pair_key: str) -> Optional[str]:
        """
        Infer a Stage-Level prefix from the first QuestionID in the pair's
        pool within the bank.  Returns None when no questions exist for the
        pair or when the bank lacks a QuestionID column.
        """
        if "QuestionID" not in self.bank.columns:
            return None
        pair_bank = self.bank[self.bank["_pair_key"] == pair_key]
        if pair_bank.empty:
            return None
        first_id = str(pair_bank["QuestionID"].iloc[0])
        return self._derive_prefix(first_id, self.config.block_prefix_segments)

    def _filter_by_prefix(self, pool: pd.DataFrame, prefix: Optional[str]) -> pd.DataFrame:
        """
        Retain only rows whose QuestionID starts with *prefix* + ".".
        If *prefix* is None, the pool is returned unchanged (no restriction).
        """
        if prefix is None or pool.empty:
            return pool
        match_prefix = prefix + "."
        mask = pool["QuestionID"].astype(str).str.startswith(match_prefix)
        # Also include exact matches (IDs that equal the prefix with no trailing dot)
        mask |= pool["QuestionID"].astype(str) == prefix
        return pool[mask]

    def _fallback_select(
        self,
        shortage: int,
        bin_spec: BinSpec,
        already_selected: set[str],
        form_idx: int,
        block_prefix: Optional[str] = None,
    ) -> tuple[list[pd.DataFrame], set[str]]:
        """
        Global fallback: pick least-used items with closest difficulty,
        ignoring pair/bin constraints.

        Used-Blocks last-resort path
        ----------------------------
        When the unused pool is exhausted and reuse becomes necessary, only
        questions whose QuestionID shares the exact same Stage-Level prefix as
        *block_prefix* are eligible.  Cross-stage or cross-level reuse is
        never allowed.  If no same-prefix questions are available in the
        Used-Blocks dataset either, the fallback returns nothing.
        """
        global_pool = self.bank[
            ~self.bank["QuestionID"].isin(already_selected)
        ].copy()

        if self.config.allow_reuse_across_forms is False:
            used = {qid for qid, cnt in self.usage_counts.items() if cnt > 0}
            global_pool = global_pool[~global_pool["QuestionID"].isin(used)]

        if global_pool.empty:
            # ── Used-Blocks last-resort path ──────────────────────────────
            # Rebuild from the full bank (including already-used questions)
            # but enforce strict Stage-Level prefix matching.
            used_blocks_pool = self.bank[
                ~self.bank["QuestionID"].isin(already_selected)
            ].copy()

            # Enforce exact Stage-Level match — no cross-stage/level reuse
            used_blocks_pool = self._filter_by_prefix(used_blocks_pool, block_prefix)

            if used_blocks_pool.empty:
                # Same-prefix pool is also exhausted; do not borrow from other
                # stages or levels — return nothing rather than violate the rule.
                if block_prefix is not None:
                    logger.warning(
                        "Form %d fallback: Used-Blocks pool for prefix '%s' is "
                        "exhausted. No cross-stage/level substitution will be made.",
                        form_idx + 1, block_prefix,
                    )
                return [], set()

            global_pool = used_blocks_pool

        if global_pool.empty:
            return [], set()

        bin_mid = (bin_spec.low + bin_spec.high) / 2.0
        dist = np.abs(global_pool["b"].to_numpy(dtype=float) - bin_mid)
        usage_vals = np.array(
            [self.usage_counts.get(str(q), 0) for q in global_pool["QuestionID"]], dtype=float
        )
        # Sort by usage ascending (least-used first), then distance ascending
        order = np.lexsort((dist, usage_vals))
        top = global_pool.iloc[order[:shortage]]
        ids = set(top["QuestionID"].astype(str).tolist())
        return [top], ids

    # ------------------------------------------------------------------
    # Pair × Bin distribution matrix
    # ------------------------------------------------------------------

    def _distribute_pairs_to_bins(
        self, pairs: list[PairSpec], bins: list[BinSpec]
    ) -> list[list[int]]:
        """
        For each pair, distribute its question count across bins
        proportionally to bank availability within each bin.
        Returns a 2D list [pair_idx][bin_idx] = count.
        """
        matrix: list[list[int]] = []
        for pair in pairs:
            pair_bank = self.bank[self.bank["_pair_key"] == pair.pair_key]
            bin_avail = []
            for b in bins:
                cnt = ((pair_bank["b"] >= b.low) & (pair_bank["b"] <= b.high)).sum()
                bin_avail.append(max(cnt, 1))  # avoid zero-division

            counts = _largest_remainder(pair.count, bin_avail)
            matrix.append(counts)
        return matrix

    # ------------------------------------------------------------------
    # Form statistics
    # ------------------------------------------------------------------

    def _compute_form_stats(self, form_df: pd.DataFrame) -> dict:
        if form_df.empty:
            return {}
        b = form_df["b"].to_numpy(dtype=float)
        a = form_df["a"].to_numpy(dtype=float)
        c = form_df["c"].to_numpy(dtype=float)
        return {
            "n_items": len(form_df),
            "mean_b": float(np.mean(b)),
            "std_b": float(np.std(b)),
            "min_b": float(np.min(b)),
            "max_b": float(np.max(b)),
            "mean_a": float(np.mean(a)),
            "mean_c": float(np.mean(c)),
        }


# ---------------------------------------------------------------------------
# Convenience factory
# ---------------------------------------------------------------------------

def build_config_from_gui(params: dict) -> AssemblyConfig:
    """
    Build an AssemblyConfig from a flat dict produced by the GUI.
    Expected keys (all optional with defaults):
        n_forms, questions_per_form, target_mean_b, target_min_b, target_max_b,
        bins               → list of {"low": float, "high": float, "proportion": float}
        pair_requirements  → list of {"pair_key": str, "الناتج": str, "المؤشر": str, "count": int}
        w_difficulty, w_bin, w_usage, w_random,
        allow_reuse_across_forms, rng_seed,
        block_prefix_segments → int (default 2); controls Stage-Level prefix depth
                                for Used-Blocks fallback matching.
    """
    bins_raw = params.get("bins", [])
    bins = [(b["low"], b["high"], b["proportion"]) for b in bins_raw]

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
        rng_seed=params.get("rng_seed", None),
        block_prefix_segments=int(params.get("block_prefix_segments", 2)),
    )
