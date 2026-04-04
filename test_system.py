"""
End-to-end integration test.
Generates a synthetic bank (500 questions, 4 pairs, 2 domains) and
assembles 4 forms of 40 questions each, then runs psychometric analysis
and writes all output files.
"""

import logging
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")


# ---------------------------------------------------------------------------
# 1. Create synthetic bank
# ---------------------------------------------------------------------------
rng = np.random.default_rng(42)
N = 500

pairs = [
    ("Outcome A", "Indicator 1"),
    ("Outcome A", "Indicator 2"),
    ("Outcome B", "Indicator 1"),
    ("Outcome B", "Indicator 3"),
]
domains = ["Domain X", "Domain Y"]

pair_idx = rng.integers(0, len(pairs), size=N)
domain_idx = rng.integers(0, len(domains), size=N)

bank_raw = pd.DataFrame({
    "QuestionID": [f"Q{i:04d}" for i in range(1, N + 1)],
    "الناتج":  [pairs[i][0] for i in pair_idx],
    "المؤشر":  [pairs[i][1] for i in pair_idx],
    "المجال":  [domains[i] for i in domain_idx],
    "b":       rng.normal(0.0, 1.0, size=N).clip(-3.5, 3.5).round(3),
    "a":       rng.uniform(0.5, 2.5, size=N).round(3),
    "c":       rng.uniform(0.05, 0.30, size=N).round(3),
})

with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
    bank_path = f.name
    bank_raw.to_csv(bank_path, index=False)

print(f"Synthetic bank saved to: {bank_path}")
print(f"Bank shape: {bank_raw.shape}")

# ---------------------------------------------------------------------------
# 2. Load and validate bank
# ---------------------------------------------------------------------------
from data_loader import load_bank, get_pair_summary

bank = load_bank(bank_path)
print(f"\nLoaded bank: {len(bank)} rows, columns: {list(bank.columns)}")
pair_summary = get_pair_summary(bank)
print("\nPair summary:")
print(pair_summary.to_string())

# ---------------------------------------------------------------------------
# 3. Assemble 4 forms of 40 questions
# ---------------------------------------------------------------------------
from assembly_engine import AssemblyConfig, AssemblyEngine

pair_reqs = []
counts = [10, 10, 10, 10]  # 40 total
for (natej, moshir), cnt in zip(pairs, counts):
    pair_key = natej + "||" + moshir
    pair_reqs.append({
        "pair_key": pair_key,
        "الناتج": natej,
        "المؤشر": moshir,
        "count": cnt,
    })

config = AssemblyConfig(
    n_forms=4,
    questions_per_form=40,
    target_mean_b=0.0,
    target_min_b=-3.0,
    target_max_b=3.0,
    bins=[
        (-3.0, -0.5, 0.25),
        (-0.5,  0.5, 0.50),
        ( 0.5,  3.0, 0.25),
    ],
    pair_requirements=pair_reqs,
    w_difficulty=0.40,
    w_bin=0.25,
    w_usage=0.20,
    w_random=0.15,
    allow_reuse_across_forms=False,
    rng_seed=42,
)

engine = AssemblyEngine(bank, config)
result = engine.assemble()

print(f"\nAssembled {len(result.forms)} forms")
if result.warnings:
    print("Warnings:")
    for w in result.warnings:
        print(f"  ⚠ {w}")

for i, (form_df, stats) in enumerate(zip(result.forms, result.stats)):
    unique_ids = form_df["QuestionID"].nunique()
    print(f"  Form {i+1}: {len(form_df)} items, {unique_ids} unique, "
          f"mean_b={stats.get('mean_b', 0):.3f}, alpha=pending")

    # Verify no duplicates within form
    assert unique_ids == len(form_df), f"Duplicate IDs in form {i+1}!"

# Check no overlap across forms (since reuse disabled)
all_ids = [set(f["QuestionID"].tolist()) for f in result.forms]
for i in range(len(all_ids)):
    for j in range(i + 1, len(all_ids)):
        overlap = all_ids[i] & all_ids[j]
        if overlap:
            print(f"  ⚠ Overlap between Form {i+1} and Form {j+1}: {len(overlap)} items")
        else:
            print(f"  ✓ No overlap between Form {i+1} and Form {j+1}")

# ---------------------------------------------------------------------------
# 4. Psychometric analysis
# ---------------------------------------------------------------------------
from psychometric import SimulationConfig, analyse_all_forms

sim_cfg = SimulationConfig(n_students=3000, theta_mean=0.0, theta_sd=1.0, rng_seed=42)
form_names = [f"Form {i+1}" for i in range(4)]
analyses = analyse_all_forms(result.forms, form_names, sim_cfg=sim_cfg)

print("\nPsychometric Summary:")
for fa in analyses:
    s = fa.summary
    print(
        f"  {fa.form_name}: alpha={s['alpha']:.3f}, "
        f"mean_b={s['mean_b']:.3f}, "
        f"mean_item_cor={s['mean_item_cor']:.3f}"
    )

# ---------------------------------------------------------------------------
# 5. Output generation
# ---------------------------------------------------------------------------
from output_generator import generate_all_outputs

out_dir = Path(tempfile.mkdtemp())
print(f"\nWriting output to: {out_dir}")
paths = generate_all_outputs(analyses, out_dir, prefix="test")

for label, p in paths.items():
    assert p.exists(), f"Output file missing: {p}"
    size_kb = p.stat().st_size / 1024
    print(f"  [{label:14s}]  {p.name}  ({size_kb:.1f} KB)")

print("\n✓ All tests passed!")
