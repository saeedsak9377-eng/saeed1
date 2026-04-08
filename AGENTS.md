# AGENTS.md

## Cursor Cloud specific instructions

This is a standalone Python desktop application for smart test assembly and psychometric analysis. There are no external services, databases, or Docker containers required.

### Services

| Service | How to run |
|---|---|
| **Integration test** | `python3 test_system.py` — self-contained, creates a synthetic bank and runs the full pipeline |
| **CLI mode** | `python3 main.py --cli --bank <path> --forms N --qpf N --output <dir>` (see `--help` for all options) |
| **GUI mode** | `python3 main.py` — requires `tkinter` and a display server (not available in headless Cloud VMs) |

### Key notes

- **Python ≥ 3.10** is required. The VM ships with Python 3.12.
- **Dependencies**: `pip install -r requirements.txt` (pandas, numpy, openpyxl, scipy, matplotlib, xlsxwriter).
- **Linting**: No project-specific linter config; use `python3 -m flake8 --max-line-length=120 *.py` for basic checks. Pre-existing style warnings exist (E221, F401, etc.) — these are in the original code.
- **Tests**: Run `python3 test_system.py` from the repo root. It creates a synthetic 500-question bank, assembles 4 forms, runs 3PL IRT analysis, and writes 3 Excel output files. Expected runtime: ~3 seconds.
- **GUI mode** will not work in headless Cloud VMs (no display server / tkinter). Use CLI mode or the integration test instead.
- **Matplotlib backend**: The codebase explicitly sets the `Agg` backend in `output_generator.py`, so chart generation works headlessly without any manual backend configuration.
