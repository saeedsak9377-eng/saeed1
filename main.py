"""
Smart Test Assembly & Psychometric Analysis
===========================================
Entry point.

Usage
-----
GUI mode (default):
    python main.py

CLI mode (headless, for automation / CI):
    python main.py --cli --bank path/to/bank.xlsx \
        --forms 4 --qpf 40 --output ./output

Use --help for full CLI reference.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# CLI helper
# ---------------------------------------------------------------------------

def _cli_main(args: argparse.Namespace) -> int:
    from data_loader import load_bank
    from assembly_engine import AssemblyConfig, AssemblyEngine
    from psychometric import SimulationConfig, analyse_all_forms
    from output_generator import generate_all_outputs

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
    )
    log = logging.getLogger("cli")

    # Load bank
    log.info("Loading bank: %s", args.bank)
    bank = load_bank(args.bank)
    log.info("  %d questions loaded", len(bank))

    # Build config
    config = AssemblyConfig(
        n_forms=args.forms,
        questions_per_form=args.qpf,
        target_mean_b=args.mean_b,
        target_min_b=args.min_b,
        target_max_b=args.max_b,
        allow_reuse_across_forms=args.reuse,
        rng_seed=args.seed,
    )

    # Assemble
    log.info("Assembling %d forms…", config.n_forms)
    engine = AssemblyEngine(bank, config)
    result = engine.assemble()
    for w in result.warnings:
        log.warning("  %s", w)

    # Analyse
    log.info("Running psychometric analysis…")
    sim_cfg = SimulationConfig(n_students=args.n_students, rng_seed=args.seed)
    form_names = [f"Form {i + 1}" for i in range(len(result.forms))]
    analyses = analyse_all_forms(result.forms, form_names, sim_cfg=sim_cfg)

    # Output
    out_dir = Path(args.output)
    log.info("Writing output to %s", out_dir)
    paths = generate_all_outputs(analyses, out_dir, prefix=args.prefix)
    for label, p in paths.items():
        log.info("  %-15s  %s", label, p)

    log.info("Done.")
    return 0


# ---------------------------------------------------------------------------
# GUI helper
# ---------------------------------------------------------------------------

def _gui_main() -> int:
    try:
        from gui_app import launch
        launch()
        return 0
    except ImportError as exc:
        print(f"GUI dependencies missing: {exc}", file=sys.stderr)
        print("Install with:  pip install -r requirements.txt", file=sys.stderr)
        return 1


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="test_assembly",
        description="Smart Test Assembly & Psychometric Analysis",
    )
    p.add_argument("--cli", action="store_true", help="Run in headless CLI mode")
    p.add_argument("--bank", type=str, help="Path to question bank (Excel/CSV)")
    p.add_argument("--forms", type=int, default=4, help="Number of forms (default 4)")
    p.add_argument("--qpf", type=int, default=40, help="Questions per form (default 40)")
    p.add_argument("--mean-b", type=float, default=0.0, dest="mean_b")
    p.add_argument("--min-b",  type=float, default=-3.0, dest="min_b")
    p.add_argument("--max-b",  type=float, default=3.0,  dest="max_b")
    p.add_argument("--reuse", action="store_true", help="Allow reuse across forms")
    p.add_argument("--seed", type=int, default=None)
    p.add_argument("--n-students", type=int, default=3000, dest="n_students")
    p.add_argument("--output", type=str, default="./output", help="Output directory")
    p.add_argument("--prefix", type=str, default="test", help="Output file prefix")
    return p


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = _build_parser()
    args = parser.parse_args()

    if args.cli:
        if not args.bank:
            parser.error("--bank is required in CLI mode")
        sys.exit(_cli_main(args))
    else:
        sys.exit(_gui_main())
