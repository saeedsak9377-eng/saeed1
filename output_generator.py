"""
Output Generator
----------------
Produces three professional Excel workbooks:

1. forms_output.xlsx
   - One sheet per form with full item table + summary header
2. global_summary.xlsx
   - Cross-form comparison of difficulty and psychometric statistics
3. psychometric_analysis.xlsx
   - ICC chart (all items) per form
   - TCC comparison across forms
   - TIF comparison across forms
   - Item correlation tables

All charts use XlsxWriter for embedded Excel charts.
Matplotlib is used only for the psychometric analysis workbook (richer styling).
"""

from __future__ import annotations

import io
import logging
from pathlib import Path
from typing import Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import numpy as np
import pandas as pd
import xlsxwriter

from psychometric import FormAnalysis

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
HEADER_BG      = "#1F3864"   # dark navy
HEADER_FONT    = "#FFFFFF"
SUBHEADER_BG   = "#2E75B6"
ALT_ROW_BG     = "#DCE6F1"
BORDER_COLOR   = "#4472C4"
CHART_COLORS   = [
    "#2E75B6", "#ED7D31", "#A9D18E", "#FF0000",
    "#7030A0", "#00B0F0", "#FF6600", "#70AD47",
]


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _add_formats(wb: xlsxwriter.Workbook) -> dict:
    """Register reusable cell formats in the workbook."""
    f = {}
    f["header"] = wb.add_format({
        "bold": True, "bg_color": HEADER_BG, "font_color": HEADER_FONT,
        "border": 1, "align": "center", "valign": "vcenter",
        "text_wrap": True,
    })
    f["subheader"] = wb.add_format({
        "bold": True, "bg_color": SUBHEADER_BG, "font_color": HEADER_FONT,
        "border": 1, "align": "center",
    })
    f["metric_label"] = wb.add_format({
        "bold": True, "bg_color": "#BDD7EE", "border": 1,
    })
    f["metric_value"] = wb.add_format({
        "bg_color": "#FFFFFF", "border": 1, "num_format": "0.0000",
    })
    f["metric_value_int"] = wb.add_format({
        "bg_color": "#FFFFFF", "border": 1, "num_format": "0",
    })
    f["data"] = wb.add_format({
        "border": 1, "align": "left", "valign": "vcenter",
    })
    f["data_num"] = wb.add_format({
        "border": 1, "align": "center", "num_format": "0.0000",
    })
    f["data_alt"] = wb.add_format({
        "bg_color": ALT_ROW_BG, "border": 1, "align": "left",
    })
    f["data_num_alt"] = wb.add_format({
        "bg_color": ALT_ROW_BG, "border": 1, "align": "center",
        "num_format": "0.0000",
    })
    f["title"] = wb.add_format({
        "bold": True, "font_size": 14, "bg_color": HEADER_BG,
        "font_color": HEADER_FONT, "align": "center", "valign": "vcenter",
    })
    return f


def _col_letters(n: int) -> str:
    """Convert 0-indexed column to Excel letter (A, B, …, Z, AA, …)."""
    result = ""
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ---------------------------------------------------------------------------
# 1. Forms output workbook
# ---------------------------------------------------------------------------

FORM_COLUMNS = [
    ("QuestionID", 18, False),
    ("الناتج",     22, False),
    ("المؤشر",     22, False),
    ("المجال",     20, False),
    ("b",          12, True),
    ("a",          12, True),
    ("c",          12, True),
]


def write_forms_workbook(
    analyses: list[FormAnalysis],
    path: Path,
) -> None:
    """Write one sheet per form with summary header and item table."""
    wb = xlsxwriter.Workbook(str(path))
    fmts = _add_formats(wb)

    for fa in analyses:
        sheet_name = fa.form_name[:31]  # Excel limit
        ws = wb.add_worksheet(sheet_name)
        ws.right_to_left()  # Arabic RTL

        _write_form_summary_header(ws, fmts, fa)
        _write_form_item_table(ws, fmts, fa)

    wb.close()
    logger.info("Forms workbook written to %s", path)


def _write_form_summary_header(ws, fmts, fa: FormAnalysis):
    s = fa.summary
    ws.merge_range("A1:D1", fa.form_name, fmts["title"])

    metrics = [
        ("Number of Items",   s.get("n_items"),      False),
        ("Mean Difficulty (b)", s.get("mean_b"),     True),
        ("Min b",             s.get("min_b"),         True),
        ("Max b",             s.get("max_b"),         True),
        ("Std b",             s.get("std_b"),         True),
        ("Mean Discrimination (a)", s.get("mean_a"), True),
        ("Mean Guessing (c)", s.get("mean_c"),        True),
        ("Cronbach Alpha",    s.get("alpha"),         True),
        ("Mean Item Corr.",   s.get("mean_item_cor"), True),
    ]
    row = 2
    ws.write(row, 0, "Metric",   fmts["subheader"])
    ws.write(row, 1, "Value",    fmts["subheader"])
    row += 1
    for label, val, is_float in metrics:
        ws.write(row, 0, label, fmts["metric_label"])
        fmt = fmts["metric_value"] if is_float else fmts["metric_value_int"]
        if val is None:
            ws.write(row, 1, "N/A", fmts["data"])
        elif is_float:
            ws.write_number(row, 1, float(val), fmt)
        else:
            ws.write_number(row, 1, int(val), fmt)
        row += 1

    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 15)
    # blank row before table
    ws._data_start_row = row + 1


def _write_form_item_table(ws, fmts, fa: FormAnalysis):
    start_row = getattr(ws, "_data_start_row", 14)
    df = fa.form_df.reset_index(drop=True)

    # Optionally include item-total correlation
    if fa.item_cors is not None and len(fa.item_cors) == len(df):
        df = df.copy()
        df["Item Corr."] = np.round(fa.item_cors, 4)

    columns = list(FORM_COLUMNS)
    if "Item Corr." in df.columns:
        columns.append(("Item Corr.", 14, True))

    # Header row
    header_row = start_row
    for col_idx, (col_name, width, _) in enumerate(columns):
        ws.write(header_row, col_idx, col_name, fmts["subheader"])
        ws.set_column(col_idx, col_idx, width)

    # Data rows
    for row_idx, (_, item_row) in enumerate(df.iterrows()):
        excel_row = header_row + 1 + row_idx
        alt = row_idx % 2 == 1
        for col_idx, (col_name, _, is_num) in enumerate(columns):
            val = item_row.get(col_name, "")
            if is_num:
                fmt = fmts["data_num_alt"] if alt else fmts["data_num"]
                try:
                    ws.write_number(excel_row, col_idx, float(val), fmt)
                except (ValueError, TypeError):
                    ws.write(excel_row, col_idx, str(val), fmt)
            else:
                fmt = fmts["data_alt"] if alt else fmts["data"]
                ws.write(excel_row, col_idx, str(val) if val is not None else "", fmt)

    ws.freeze_panes(header_row + 1, 0)
    ws.autofilter(header_row, 0, header_row + len(df), len(columns) - 1)


# ---------------------------------------------------------------------------
# 2. Global summary workbook
# ---------------------------------------------------------------------------

def write_global_summary(
    analyses: list[FormAnalysis],
    path: Path,
) -> None:
    wb = xlsxwriter.Workbook(str(path))
    fmts = _add_formats(wb)
    ws = wb.add_worksheet("Cross-Form Summary")

    cols = [
        "Form", "n_items", "mean_b", "std_b", "min_b", "max_b",
        "mean_a", "mean_c", "alpha", "mean_item_cor",
        "min_item_cor", "max_item_cor",
        "peak_info_theta", "peak_info",
    ]
    header_labels = {
        "Form": "Form", "n_items": "N Items",
        "mean_b": "Mean b", "std_b": "SD b",
        "min_b": "Min b", "max_b": "Max b",
        "mean_a": "Mean a", "mean_c": "Mean c",
        "alpha": "Alpha", "mean_item_cor": "Mean Item Corr.",
        "min_item_cor": "Min Item Corr.", "max_item_cor": "Max Item Corr.",
        "peak_info_theta": "Peak Info θ", "peak_info": "Peak Info",
    }
    widths = [18, 9, 10, 10, 10, 10, 10, 10, 10, 14, 14, 14, 14, 12]

    ws.merge_range(0, 0, 0, len(cols) - 1, "Cross-Form Psychometric Summary", fmts["title"])
    ws.set_row(0, 22)

    for ci, (col, w) in enumerate(zip(cols, widths)):
        ws.write(1, ci, header_labels[col], fmts["subheader"])
        ws.set_column(ci, ci, w)

    for ri, fa in enumerate(analyses):
        s = fa.summary
        row_data = [
            fa.form_name, s.get("n_items"), s.get("mean_b"), s.get("std_b"),
            s.get("min_b"), s.get("max_b"), s.get("mean_a"), s.get("mean_c"),
            s.get("alpha"), s.get("mean_item_cor"), s.get("min_item_cor"),
            s.get("max_item_cor"), s.get("peak_info_theta"), s.get("peak_info"),
        ]
        alt = ri % 2 == 0
        for ci, val in enumerate(row_data):
            fmt = fmts["data_alt"] if alt else fmts["data"]
            fmt_n = fmts["data_num_alt"] if alt else fmts["data_num"]
            if ci == 0:
                ws.write(2 + ri, ci, str(val), fmt)
            elif val is None:
                ws.write(2 + ri, ci, "N/A", fmt)
            elif ci == 1:
                ws.write_number(2 + ri, ci, int(val), fmt)
            else:
                ws.write_number(2 + ri, ci, float(val), fmt_n)

    ws.freeze_panes(2, 0)
    wb.close()
    logger.info("Global summary written to %s", path)


# ---------------------------------------------------------------------------
# 3. Psychometric analysis workbook (Matplotlib charts embedded as images)
# ---------------------------------------------------------------------------

def write_psychometric_workbook(
    analyses: list[FormAnalysis],
    path: Path,
) -> None:
    wb = xlsxwriter.Workbook(str(path))
    fmts = _add_formats(wb)

    for fa in analyses:
        _write_icc_sheet(wb, fmts, fa)

    _write_tcc_sheet(wb, fmts, analyses)
    _write_tif_sheet(wb, fmts, analyses)
    _write_item_corr_sheet(wb, fmts, analyses)

    wb.close()
    logger.info("Psychometric workbook written to %s", path)


# ---- ICC sheet ----
def _write_icc_sheet(wb, fmts, fa: FormAnalysis):
    sheet_name = f"ICC - {fa.form_name}"[:31]
    ws = wb.add_worksheet(sheet_name)
    ws.merge_range("A1:B1", f"ICC – {fa.form_name}", fmts["title"])

    fig, ax = plt.subplots(figsize=(10, 6))
    K = fa.P_grid.shape[1]
    cmap = cm.get_cmap("tab20", K)
    for j in range(K):
        qid = (
            fa.form_df["QuestionID"].iloc[j]
            if j < len(fa.form_df)
            else f"Item {j + 1}"
        )
        ax.plot(
            fa.theta_grid, fa.P_grid[:, j],
            color=cmap(j), linewidth=1.2, alpha=0.85,
            label=str(qid),
        )
    ax.set_xlabel("Ability (θ)", fontsize=12)
    ax.set_ylabel("P(correct)", fontsize=12)
    ax.set_title(f"Item Characteristic Curves – {fa.form_name}", fontsize=13, pad=10)
    ax.set_xlim(-4, 4)
    ax.set_ylim(0, 1.05)
    ax.grid(True, alpha=0.3)
    if K <= 20:
        ax.legend(
            loc="upper left", fontsize=7,
            ncol=2, framealpha=0.7,
        )
    fig.tight_layout()
    img_bytes = _fig_to_bytes(fig)
    plt.close(fig)
    ws.insert_image("D2", "", {"image_data": img_bytes, "x_scale": 1.0, "y_scale": 1.0})


# ---- TCC sheet ----
def _write_tcc_sheet(wb, fmts, analyses: list[FormAnalysis]):
    ws = wb.add_worksheet("TCC")
    ws.merge_range("A1:B1", "Test Characteristic Curves", fmts["title"])

    fig, ax = plt.subplots(figsize=(10, 5))
    for i, fa in enumerate(analyses):
        color = CHART_COLORS[i % len(CHART_COLORS)]
        ax.plot(fa.theta_grid, fa.TCC, color=color, linewidth=2.0, label=fa.form_name)
    ax.set_xlabel("Ability (θ)", fontsize=12)
    ax.set_ylabel("Expected Score", fontsize=12)
    ax.set_title("Test Characteristic Curves", fontsize=13, pad=10)
    ax.set_xlim(-4, 4)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left", fontsize=10)
    fig.tight_layout()
    img_bytes = _fig_to_bytes(fig)
    plt.close(fig)
    ws.insert_image("D2", "", {"image_data": img_bytes})


# ---- TIF sheet ----
def _write_tif_sheet(wb, fmts, analyses: list[FormAnalysis]):
    ws = wb.add_worksheet("TIF")
    ws.merge_range("A1:B1", "Test Information Functions", fmts["title"])

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))

    # TIF
    for i, fa in enumerate(analyses):
        color = CHART_COLORS[i % len(CHART_COLORS)]
        axes[0].plot(fa.theta_grid, fa.TIF, color=color, linewidth=2.0, label=fa.form_name)
    axes[0].set_xlabel("Ability (θ)", fontsize=12)
    axes[0].set_ylabel("Information", fontsize=12)
    axes[0].set_title("Test Information Function", fontsize=13, pad=8)
    axes[0].set_xlim(-4, 4)
    axes[0].grid(True, alpha=0.3)
    axes[0].legend(fontsize=9)

    # SEM
    for i, fa in enumerate(analyses):
        color = CHART_COLORS[i % len(CHART_COLORS)]
        axes[1].plot(fa.theta_grid, fa.SEM, color=color, linewidth=2.0, label=fa.form_name)
    axes[1].set_xlabel("Ability (θ)", fontsize=12)
    axes[1].set_ylabel("SEM", fontsize=12)
    axes[1].set_title("Standard Error of Measurement", fontsize=13, pad=8)
    axes[1].set_xlim(-4, 4)
    axes[1].set_ylim(0, 2.0)
    axes[1].grid(True, alpha=0.3)
    axes[1].legend(fontsize=9)

    fig.tight_layout()
    img_bytes = _fig_to_bytes(fig)
    plt.close(fig)
    ws.insert_image("D2", "", {"image_data": img_bytes})


# ---- Item correlation sheet ----
def _write_item_corr_sheet(wb, fmts, analyses: list[FormAnalysis]):
    ws = wb.add_worksheet("Item Correlations")
    ws.merge_range("A1:C1", "Corrected Item-Total Correlations", fmts["title"])

    row = 2
    for fa in analyses:
        ws.write(row, 0, fa.form_name, fmts["subheader"])
        ws.write(row, 1, "QuestionID", fmts["subheader"])
        ws.write(row, 2, "Item-Total Corr.", fmts["subheader"])
        ws.write(row, 3, "b (difficulty)", fmts["subheader"])
        row += 1

        K = len(fa.form_df)
        for j in range(K):
            qid = fa.form_df["QuestionID"].iloc[j] if j < len(fa.form_df) else f"Item{j + 1}"
            b_val = fa.form_df["b"].iloc[j] if j < len(fa.form_df) else float("nan")
            cor = fa.item_cors[j] if fa.item_cors is not None and j < len(fa.item_cors) else float("nan")
            alt = j % 2 == 0
            fmt = fmts["data_alt"] if alt else fmts["data"]
            fmt_n = fmts["data_num_alt"] if alt else fmts["data_num"]
            ws.write(row + j, 0, str(j + 1), fmt)
            ws.write(row + j, 1, str(qid), fmt)
            ws.write_number(row + j, 2, float(cor) if not np.isnan(cor) else 0.0, fmt_n)
            ws.write_number(row + j, 3, float(b_val), fmt_n)

        row += K + 1

    ws.set_column(0, 0, 8)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 15)

    # Embed correlation distribution chart
    _write_item_corr_chart(wb, ws, analyses, start_col=5)


def _write_item_corr_chart(wb, ws, analyses, start_col=5):
    """Small box-plot style figure showing item-correlation distributions."""
    fig, ax = plt.subplots(figsize=(8, 4))
    data = [fa.item_cors[~np.isnan(fa.item_cors)] for fa in analyses if fa.item_cors is not None]
    labels = [fa.form_name for fa in analyses if fa.item_cors is not None]
    if data:
        bp = ax.boxplot(data, labels=labels, patch_artist=True, notch=False)
        for patch, color in zip(bp["boxes"], CHART_COLORS):
            patch.set_facecolor(color)
            patch.set_alpha(0.7)
    ax.set_ylabel("Item-Total Correlation")
    ax.set_title("Item Correlation Distribution by Form")
    ax.grid(True, alpha=0.3, axis="y")
    fig.tight_layout()
    img_bytes = _fig_to_bytes(fig)
    plt.close(fig)
    ws.insert_image(2, start_col, "", {"image_data": img_bytes})


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def _fig_to_bytes(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Master output function
# ---------------------------------------------------------------------------

def generate_all_outputs(
    analyses: list[FormAnalysis],
    output_dir: Path,
    prefix: str = "test",
) -> dict[str, Path]:
    """
    Write all three output workbooks to output_dir.
    Returns a dict mapping label → Path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "forms":         output_dir / f"{prefix}_forms.xlsx",
        "summary":       output_dir / f"{prefix}_global_summary.xlsx",
        "psychometric":  output_dir / f"{prefix}_psychometric.xlsx",
    }
    write_forms_workbook(analyses, paths["forms"])
    write_global_summary(analyses, paths["summary"])
    write_psychometric_workbook(analyses, paths["psychometric"])
    return paths
