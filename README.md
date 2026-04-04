# Smart Test Assembly & Psychometric Analysis

A production-quality Python system for assembling parallel test forms from a
question bank and automatically analysing them using full 3PL IRT psychometrics.

---

## Features

### Intelligent Assembly Engine
- **Scoring-based selection** — every item receives a composite score:
  ```
  score = w_difficulty × difficulty_fit
        + w_bin        × bin_fit
        + w_usage      × usage_penalty
        + w_random     × noise
  ```
  Items are selected greedily (highest score first) with a small randomness
  component to produce variation across forms without sacrificing quality.
- **Largest Remainder Method** for exact integer bin quotas (zero rounding drift).
- **(الناتج, المؤشر) pair handling** — each (outcome, indicator) pair is treated
  as an independent entity, even when indicator values are numerically identical
  under different outcomes.
- **No duplicate QuestionIDs** within a form; strong minimisation across forms.
- **Adaptive shortage fallback** — four relaxation stages instead of placeholders:
  1. Strict bin membership
  2. Relaxed bin edges (±20 % of half-width)
  3. Closest-difficulty from full pair pool
  4. Global least-used / closest-difficulty fallback
- Handles banks of 10 000+ questions using vectorised NumPy/Pandas operations.

### Psychometric Engine (3PL IRT)
- Full vectorised 3PL model: `P(θ) = c + (1−c) / (1 + exp(−a(θ−b)))`
- Simulates **3 000+ students** by default (configurable)
- Computes per-form:
  - **ICC** (Item Characteristic Curves) — one line per item
  - **TCC** (Test Characteristic Curve)
  - **TIF** (Test Information Function) + SEM
  - **Cronbach's Alpha** (KR-20)
  - **Corrected item-total correlations**
- Cross-form comparison table

### Professional Excel Output
Three workbooks are written automatically after assembly:

| File | Contents |
|---|---|
| `{prefix}_forms.xlsx` | One sheet per form: summary header + full item table |
| `{prefix}_global_summary.xlsx` | Cross-form difficulty & psychometric comparison |
| `{prefix}_psychometric.xlsx` | ICC/TCC/TIF charts + item correlation tables |

### Smart GUI
- Four-tab workflow: **Bank → Configure → Generate → Results**
- Grid-based bin editor (add/remove rows dynamically)
- Auto-populated (الناتج, المؤشر) pair distribution table
- Background threading — GUI never freezes
- Live progress bar + colour-coded log console

---

## Installation

```bash
pip install -r requirements.txt
```

Requires Python ≥ 3.10.

---

## Usage

### GUI (default)
```bash
python3 main.py
```

### CLI (headless / automation)
```bash
python3 main.py --cli \
    --bank path/to/bank.xlsx \
    --forms 4 \
    --qpf 40 \
    --mean-b 0.0 \
    --min-b -3.0 \
    --max-b 3.0 \
    --n-students 3000 \
    --output ./output \
    --prefix mytest
```

---

## Question Bank Format

Excel (`.xlsx`) or CSV (`.csv`) with the following columns
(Arabic or English names accepted):

| Canonical | Aliases |
|---|---|
| `QuestionID` | `question_id`, `id`, `رقم السؤال` |
| `الناتج` | `outcome`, `learning_outcome` |
| `المؤشر` | `indicator` |
| `المجال` | `domain` |
| `b` | `difficulty`, `صعوبة` |
| `a` | `تمييز`, `discrimination` |
| `c` | `تخمين`, `guessing` |

Missing IRT parameters are filled with safe defaults (`a=1.0`, `b=0.0`, `c=0.25`).
Out-of-range values are clamped. Duplicate `QuestionID` rows are deduplicated
with a warning logged.

---

## Architecture

```
main.py                 ← Entry point (GUI or CLI)
gui_app.py              ← Tkinter GUI (4-tab workflow)
data_loader.py          ← Bank loading, column mapping, validation
assembly_engine.py      ← Scoring-based greedy assembly engine
psychometric.py         ← 3PL IRT, simulation, ICC/TCC/TIF, alpha
output_generator.py     ← Excel workbooks + embedded Matplotlib charts
requirements.txt        ← Python dependencies
test_system.py          ← Integration test (synthetic bank)
```

---

## Running Tests

```bash
python3 test_system.py
```

Expected output: 4 forms assembled with no overlaps, alpha ≈ 0.87–0.90,
three output Excel files written.
