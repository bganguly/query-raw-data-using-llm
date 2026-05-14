#!/usr/bin/env python3
"""Generate PNG preview images for each pipeline stage and save to docs/images/.

Stages:
  1. remote-source  — first rows of the raw DOL XLSX (fetched from DOL.gov)
  2. normalized-csv — first rows of the local combined CSV
  3. parquet        — first rows of the local single-file parquet

Usage:
    python3 scripts/generate_data_previews.py [--fy FY] [--quarter Q] [--rows N]
    python3 scripts/generate_data_previews.py --skip-remote   # offline / no download
"""

import argparse
import json
import pathlib
import subprocess
import sys
import tempfile

import pyarrow.csv as pa_csv
import pyarrow.parquet as pq

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except ImportError:
    print("matplotlib is required.  pip install matplotlib", file=sys.stderr)
    sys.exit(1)

try:
    from openpyxl import load_workbook
except ImportError:
    print("openpyxl is required.  pip install openpyxl", file=sys.stderr)
    sys.exit(1)

ROOT = pathlib.Path(__file__).resolve().parents[1]
DOCS_IMAGES = ROOT / "docs" / "images"

DOL_XLSX_URL = (
    "https://www.dol.gov/sites/dolgov/files/ETA/oflc/pdfs/"
    "LCA_Disclosure_Data_FY{fy}_Q{quarter}.xlsx"
)

# Representative subset of raw DOL XLSX columns to display
XLSX_DISPLAY_COLS = [
    "EMPLOYER_NAME",
    "JOB_TITLE",
    "WORKSITE_CITY",
    "WORKSITE_STATE",
    "WAGE_RATE_OF_PAY_FROM",
    "CASE_STATUS",
    "CASE_SUBMITTED",
]

# Columns to display from the normalized CSV / parquet
NORMALIZED_DISPLAY_COLS = [
    "employer",
    "job_title",
    "work_location",
    "wage",
    "status",
    "year",
    "fiscal_year",
    "fiscal_quarter",
]

HEADER_COLOR = "#2c7be5"
ROW_COLORS = ["#eef3fb", "#ffffff"]
HEADER_TEXT_COLOR = "white"
MAX_CELL_LEN = 28  # truncate long values so columns stay narrow


def _truncate(value) -> str:
    text = str(value) if value is not None else "—"
    return text[:MAX_CELL_LEN] + "…" if len(text) > MAX_CELL_LEN else text


def render_table_png(
    headers: list[str],
    rows: list[list],
    title: str,
    output_path: pathlib.Path,
    col_width: float = 2.2,
) -> None:
    """Render a list of rows as a styled PNG table using matplotlib."""
    n_rows = len(rows)
    n_cols = len(headers)

    fig_w = max(10, col_width * n_cols)
    fig_h = 0.5 * (n_rows + 1) + 1.0

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")

    display_rows = [[_truncate(v) for v in row] for row in rows]

    tbl = ax.table(
        cellText=display_rows,
        colLabels=headers,
        cellLoc="left",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)
    tbl.auto_set_column_width(list(range(n_cols)))
    tbl.scale(1, 1.4)

    for j in range(n_cols):
        cell = tbl[0, j]
        cell.set_facecolor(HEADER_COLOR)
        cell.set_text_props(color=HEADER_TEXT_COLOR, fontweight="bold")
        cell.set_edgecolor("#1a5fbc")

    for i in range(1, n_rows + 1):
        for j in range(n_cols):
            cell = tbl[i, j]
            cell.set_facecolor(ROW_COLORS[(i - 1) % 2])
            cell.set_edgecolor("#cccccc")

    fig.suptitle(title, fontsize=9, fontweight="bold", y=0.98)
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  Saved: {output_path.relative_to(ROOT)}")


# ---------------------------------------------------------------------------
# Stage helpers
# ---------------------------------------------------------------------------

def fetch_xlsx_preview(fy: int, quarter: int, n_rows: int) -> tuple[list[str], list[list]]:
    """Download the DOL XLSX and return (headers, rows) for the display columns."""
    url = DOL_XLSX_URL.format(fy=fy, quarter=quarter)
    print(f"Downloading {url} …")
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = pathlib.Path(tmp.name)

    result = subprocess.run(
        ["curl", "-L", "--fail", "-s", url, "-o", str(tmp_path)],
        capture_output=True,
    )
    if result.returncode != 0:
        tmp_path.unlink(missing_ok=True)
        raise RuntimeError(f"curl failed for {url}")

    wb = load_workbook(filename=tmp_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows_iter = ws.iter_rows(values_only=True)
    raw_header = [str(v).strip() if v is not None else "" for v in next(rows_iter)]

    header_upper = [h.upper() for h in raw_header]
    col_indices: list[int] = []
    found_headers: list[str] = []
    for col in XLSX_DISPLAY_COLS:
        try:
            idx = header_upper.index(col.upper())
            col_indices.append(idx)
            found_headers.append(col)
        except ValueError:
            pass  # column absent in this FY schema

    data_rows: list[list] = []
    for i, row in enumerate(rows_iter):
        if i >= n_rows:
            break
        data_rows.append([row[idx] if idx < len(row) else None for idx in col_indices])

    wb.close()
    tmp_path.unlink(missing_ok=True)
    return found_headers, data_rows


def read_csv_preview(csv_path: pathlib.Path, n_rows: int) -> tuple[list[str], list[list]]:
    table = pa_csv.read_csv(csv_path)
    table = table.slice(0, n_rows)
    available = [c for c in NORMALIZED_DISPLAY_COLS if c in table.schema.names]
    table = table.select(available)
    headers = list(table.schema.names)
    rows = [[table[col][i].as_py() for col in headers] for i in range(table.num_rows)]
    return headers, rows


def read_parquet_preview(parquet_path: pathlib.Path, n_rows: int) -> tuple[list[str], list[list]]:
    available = [
        c for c in NORMALIZED_DISPLAY_COLS
        if c in pq.read_schema(parquet_path).names
    ]
    table = pq.read_table(parquet_path, columns=available)
    table = table.slice(0, n_rows)
    headers = list(table.schema.names)
    rows = [[table[col][i].as_py() for col in headers] for i in range(table.num_rows)]
    return headers, rows


def load_manifest() -> dict:
    manifest_path = ROOT / "data" / "manifest.json"
    if manifest_path.exists():
        return json.loads(manifest_path.read_text())
    return {}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate data-preview PNGs for the README."
    )
    parser.add_argument(
        "--fy", type=int, default=None,
        help="Fiscal year for the remote XLSX preview (default: from manifest or 2026).",
    )
    parser.add_argument(
        "--quarter", type=int, default=None,
        help="Quarter for the remote XLSX preview (default: from manifest or 1).",
    )
    parser.add_argument(
        "--rows", type=int, default=5,
        help="Number of preview rows per stage (default: 5).",
    )
    parser.add_argument(
        "--skip-remote", action="store_true",
        help="Skip the remote XLSX download (useful when offline).",
    )
    args = parser.parse_args()

    manifest = load_manifest()
    fy = args.fy or manifest.get("last_fy") or 2026
    quarter = args.quarter or manifest.get("last_quarter") or 1

    DOCS_IMAGES.mkdir(parents=True, exist_ok=True)

    # ── Stage 1: Remote DOL XLSX ──────────────────────────────────────────
    if not args.skip_remote:
        print(f"\n[1/3] Remote XLSX preview  (FY{fy} Q{quarter})")
        try:
            headers, rows = fetch_xlsx_preview(fy, quarter, args.rows)
            render_table_png(
                headers, rows,
                f"Remote Source — DOL XLSX  FY{fy} Q{quarter}  "
                f"(first {args.rows} rows, selected columns)",
                DOCS_IMAGES / "preview-remote-xlsx.png",
                col_width=1.9,
            )
        except Exception as exc:
            print(f"  Warning: {exc}")
    else:
        print("\n[1/3] Skipping remote XLSX preview (--skip-remote).")

    # ── Stage 2: Normalized CSV ───────────────────────────────────────────
    csv_path = ROOT / "data" / "dol_lca_h1b_combined.csv"
    print("\n[2/3] Normalized CSV preview")
    if csv_path.exists():
        headers, rows = read_csv_preview(csv_path, args.rows)
        render_table_png(
            headers, rows,
            f"Local Normalized CSV  (first {args.rows} rows)",
            DOCS_IMAGES / "preview-normalized-csv.png",
            col_width=2.2,
        )
    else:
        print(f"  CSV not found: {csv_path.relative_to(ROOT)} — skipping.")

    # ── Stage 3: Parquet ─────────────────────────────────────────────────
    parquet_path = ROOT / "data" / "parquet" / "dol_lca_h1b_combined.parquet"
    print("\n[3/3] Parquet preview")
    if parquet_path.exists():
        headers, rows = read_parquet_preview(parquet_path, args.rows)
        render_table_png(
            headers, rows,
            f"Local Parquet  (first {args.rows} rows)",
            DOCS_IMAGES / "preview-parquet.png",
            col_width=2.2,
        )
    else:
        print(f"  Parquet not found: {parquet_path.relative_to(ROOT)} — skipping.")

    print("\nDone.")


if __name__ == "__main__":
    main()
