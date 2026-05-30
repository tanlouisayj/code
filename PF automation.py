"""
PF Caselist Automation Pipeline
================================
Drop a YYYYMM.xlsx file into the watch folder and this script will:
  1. Filter rows for that YYYYMM from the source file
  2. Replace / append those rows in 'QS SQL download.xlsx' > 'QS data' tab
  3. Update AE / AF formulas dynamically
  4. Refresh all pivots in 'Caselist Pivots' tab
  5. Copy 5 pivots into a new output file
  6. Calculate MoM NII & Fee movement, find top contributors (≥80% of move)
  7. Send email to Steve via Outlook with attachment + MoM commentary

REQUIREMENTS (install once):
    pip install watchdog xlwings pandas openpyxl pywin32

Run this script on your Windows PC / keep it running in the background.
"""

import os
import re
import time
import logging
import calendar

import pandas as pd
import xlwings as xw
import win32com.client
from openpyxl.utils import get_column_letter, column_index_from_string
from watchdog.observers.polling import PollingObserver   # PollingObserver works on UNC network paths
from watchdog.events import FileSystemEventHandler

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION  ← edit these paths before running
# ══════════════════════════════════════════════════════════════════════════════

WATCH_FOLDER = r"\\w01g1bnkfps02y\CIB_BUC_VOL1\CIB-BUC\Project Finance\2026\Automation"

QS_FILE      = r"\\w01g1bnkfps02y\CIB_BUC_VOL1\CIB-BUC\Project Finance\2026\Automation\QS SQL download.xlsx"

# Output file will be saved here (same as watch folder unless you change it)
OUTPUT_FOLDER = WATCH_FOLDER

# Email recipients
EMAIL_TO  = "stevelim@dbs.com"
EMAIL_CC  = "kaiyuanwong@dbs.com; ganeshp@dbs.com"
EMAIL_FROM_NAME = "Louisa"

# ── Column mapping in 'Check duplicate EGs' ──────────────────────────────────
# Column Q = 202601, R = 202602, S = 202603, T = 202604, U = 202605, V = 202606 ...
# TODO: confirm BASE_MONTH below matches the month that maps to column Q
BASE_MONTH   = 202601
BASE_COL_LTR = "Q"

# ── Sheet names ───────────────────────────────────────────────────────────────
QS_DATA_TAB         = "QS data"
CHECK_DUP_TAB       = "Check duplicate EGs"
CASELIST_PIVOTS_TAB = "Caselist Pivots"

# ── Column positions in QS data (1-indexed, A=1) ─────────────────────────────
COL_AA = 27   # YYYYMM filter column
COL_AE = 31   # concatenation formula
COL_AF = 32   # dynamic vlookup formula
COL_Q  = 17   # product/loan type (checked for "card")
COL_F  = 6
COL_E  = 5
COL_N  = 14

# ── Pivot positions in Caselist Pivots tab → target sheet names in output ────
PIVOT_MAP = [
    ("A1",  "EOP Asset Balance"),
    ("S1",  "Total Income"),
    ("AP1", "NII"),
    ("BH1", "Fee"),
    ("BZ1", "Other income"),
]
NII_PIVOT_IDX = 2   # AP1
FEE_PIVOT_IDX = 3   # BH1

# ══════════════════════════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def yyyymm_to_col_letter(yyyymm: int) -> str:
    """Return the lookup column letter in Check duplicate EGs for a given YYYYMM."""
    base_yr, base_mo = BASE_MONTH // 100, BASE_MONTH % 100
    curr_yr, curr_mo = yyyymm // 100, yyyymm % 100
    diff = (curr_yr - base_yr) * 12 + (curr_mo - base_mo)
    idx  = column_index_from_string(BASE_COL_LTR) + diff
    return get_column_letter(idx)


def prev_yyyymm(yyyymm: int) -> int:
    yr, mo = yyyymm // 100, yyyymm % 100
    if mo == 1:
        return (yr - 1) * 100 + 12
    return yr * 100 + (mo - 1)


def month_label(yyyymm: int) -> str:
    """202605 → 'May 26'"""
    yr, mo = yyyymm // 100, yyyymm % 100
    return f"{calendar.month_abbr[mo]} {str(yr)[2:]}"


# ══════════════════════════════════════════════════════════════════════════════
# DATA COMPLETENESS CHECKS
# ══════════════════════════════════════════════════════════════════════════════

def run_checks(df: pd.DataFrame, yyyymm: int) -> list[str]:
    """Return a list of error messages. Empty list = all checks passed."""
    errors = []

    # Check 1: minimum 30 columns
    if df.shape[1] < 30:
        errors.append(f"Source file has only {df.shape[1]} columns — expected at least 30.")

    # Check 2: YYYYMM column exists and has data for the target period
    if df.shape[1] >= COL_AA:
        col_aa = df.iloc[:, COL_AA - 1]
        if not (col_aa == yyyymm).any():
            errors.append(f"No rows found where column AA = {yyyymm}.")
    else:
        errors.append("Column AA (column 27) does not exist in source file.")

    # Check 3: no completely blank rows in the filtered set
    filtered = df[df.iloc[:, COL_AA - 1] == yyyymm]
    blank_rows = filtered.isnull().all(axis=1).sum()
    if blank_rows > 0:
        errors.append(f"{blank_rows} completely blank rows found in filtered data.")

    # Check 4: key columns not all null
    for col_idx, col_name in [(COL_E - 1, "E"), (COL_F - 1, "F"), (COL_N - 1, "N")]:
        if filtered.iloc[:, col_idx].isnull().all():
            errors.append(f"Column {col_name} is entirely blank in the filtered data.")

    return errors


# ══════════════════════════════════════════════════════════════════════════════
# MOM ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════

def get_mom_contributors(pivot_data: list, label: str, yyyymm: int) -> str:
    """
    From a pivot table (list of lists, header in row 0),
    find contributors explaining ≥80% of MoM movement for the given YYYYMM.
    """
    if not pivot_data or len(pivot_data) < 2:
        return f"{label}: no data available for MoM analysis."

    header = pivot_data[0]
    prev_mm = prev_yyyymm(yyyymm)

    # Locate current and previous month columns
    curr_col = next((i for i, h in enumerate(header) if h == yyyymm), None)
    prev_col = next((i for i, h in enumerate(header) if h == prev_mm), None)

    if curr_col is None or prev_col is None:
        return (f"{label}: could not locate columns for {yyyymm} or {prev_mm} "
                f"in pivot header — MoM commentary skipped.")

    # Calculate row-level movement (skip header + Grand Total rows)
    movements = []
    for row in pivot_data[1:]:
        if not row or not row[0]:
            continue
        label_val = str(row[0]).strip()
        if label_val.lower() in ("grand total", "total", ""):
            continue
        curr_val = row[curr_col] if row[curr_col] is not None else 0
        prev_val = row[prev_col] if row[prev_col] is not None else 0
        movement = (curr_val or 0) - (prev_val or 0)
        if movement != 0:
            movements.append((label_val, movement))

    if not movements:
        return f"{label}: no MoM movement detected."

    movements.sort(key=lambda x: abs(x[1]), reverse=True)
    total_move = sum(m for _, m in movements)

    # Accumulate contributors until ≥80% of absolute total is explained
    contributors = []
    cumulative   = 0
    abs_total    = sum(abs(m) for _, m in movements)

    for name, mov in movements:
        contributors.append((name, mov))
        cumulative += abs(mov)
        if abs_total > 0 and cumulative / abs_total >= 0.80:
            break

    # Format output
    curr_lbl = month_label(yyyymm)
    prev_lbl = month_label(prev_mm)
    direction = "up" if total_move >= 0 else "down"

    lines = [
        f"{label} MoM ({prev_lbl} → {curr_lbl}): "
        f"{abs(total_move):,.0f} {direction}",
        f"Top contributors (explaining ≥80% of movement):",
    ]
    for name, mov in contributors:
        pct = (mov / total_move * 100) if total_move else 0
        arrow = "▲" if mov >= 0 else "▼"
        lines.append(f"  {arrow} {name}: {mov:+,.0f}  ({pct:+.1f}%)")

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL
# ══════════════════════════════════════════════════════════════════════════════

def send_email(attachment_path: str, yyyymm: int,
               nii_commentary: str, fee_commentary: str) -> None:
    period = month_label(yyyymm)
    subject = f"PF {period} Caselist - EOP Asset Balance and Total Income by Country"

    body = f"""Hi Steve,

Please find attached the updated PF Caselist - EOP Asset Balance and Total Income by Country.

{nii_commentary}

{fee_commentary}

Best regards,
{EMAIL_FROM_NAME}"""

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To      = EMAIL_TO
    mail.CC      = EMAIL_CC
    mail.Subject = subject
    mail.Body    = body
    mail.Attachments.Add(attachment_path)
    mail.Send()
    log.info(f"Email sent to {EMAIL_TO} (CC: {EMAIL_CC})")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def process(src_path: str, yyyymm: int) -> None:
    log.info(f"── Starting pipeline for period {yyyymm} ──")

    # ── Step 1: read & validate source file ───────────────────────────────────
    log.info("Reading source file …")
    df_src = pd.read_excel(src_path, header=0)

    errors = run_checks(df_src, yyyymm)
    if errors:
        log.error("Data completeness checks FAILED:")
        for e in errors:
            log.error(f"  ✗ {e}")
        log.error("Pipeline aborted — please fix the source file and re-drop it.")
        return

    log.info("All data checks passed ✓")

    # Filter rows for this YYYYMM (columns A–AD only, 30 cols)
    df_filtered = df_src[df_src.iloc[:, COL_AA - 1] == yyyymm].copy()
    data_to_write = df_filtered.iloc[:, :30].values.tolist()
    log.info(f"Filtered {len(data_to_write)} rows for {yyyymm}")

    # Lookup column for this month in Check duplicate EGs
    lookup_col = yyyymm_to_col_letter(yyyymm)
    log.info(f"VLOOKUP column for {yyyymm}: {lookup_col}")

    # ── Steps 2-6: Excel operations via xlwings ────────────────────────────────
    app = xw.App(visible=False)
    app.display_alerts  = False
    app.screen_updating = False

    nii_data = None
    fee_data = None

    try:
        wb = app.books.open(QS_FILE)
        qs  = wb.sheets[QS_DATA_TAB]

        # ── Step 2: delete existing rows for this YYYYMM ──────────────────────
        log.info("Removing existing rows for this period …")
        last_row = qs.range("A1").end("down").row

        # Read column AA only for speed
        aa_vals = qs.range(f"AA2:AA{last_row}").value
        if not isinstance(aa_vals, list):
            aa_vals = [aa_vals]

        rows_to_delete = [
            i + 2 for i, v in enumerate(aa_vals)
            if v == yyyymm or int(v) == yyyymm if v is not None
        ]

        for row_num in reversed(rows_to_delete):
            qs.rows[row_num - 1].delete()

        log.info(f"Deleted {len(rows_to_delete)} existing rows")

        # ── Step 3: append new data (cols A–AD) ───────────────────────────────
        new_last = qs.range("A1").end("down").row + 1
        qs.range(f"A{new_last}").value = data_to_write
        log.info(f"Pasted {len(data_to_write)} new rows starting at row {new_last}")

        # ── Step 4: write AE / AF formulas for new rows ───────────────────────
        final_row = new_last + len(data_to_write) - 1

        for r in range(new_last, final_row + 1):
            # AE: concatenate F, E, N
            qs.range(f"AE{r}").formula = f"=F{r}&E{r}&N{r}"

            # AF: if Q contains "card" → "Card" (not *No)
            #     else VLOOKUP in dynamic column of Check duplicate EGs
            af = (
                f'=IF(ISNUMBER(SEARCH("card",Q{r})),"Card",'
                f'IFERROR(VLOOKUP(AE{r},\'{CHECK_DUP_TAB}\'!{lookup_col}:{lookup_col},1,0),"*No"))'
            )
            qs.range(f"AF{r}").formula = af

        log.info("AE / AF formulas written ✓")

        # ── Step 5: refresh all pivots ─────────────────────────────────────────
        log.info("Refreshing all pivots …")
        wb.api.RefreshAll()
        time.sleep(5)   # give Excel time to finish refreshing
        log.info("Pivots refreshed ✓")

        # ── Step 6: copy pivots to output file ────────────────────────────────
        pivot_sheet = wb.sheets[CASELIST_PIVOTS_TAB]
        period_lbl  = month_label(yyyymm)
        out_name    = (f"PF {period_lbl} Caselist - "
                       f"EOP Asset Balance and Total Income by Country.xlsx")
        out_path    = os.path.join(OUTPUT_FOLDER, out_name)

        out_app = xw.App(visible=False)
        out_wb  = out_app.books.add()

        # Remove extra default sheets
        while len(out_wb.sheets) > 1:
            out_wb.sheets[-1].delete()

        for idx, (pos, tab_name) in enumerate(PIVOT_MAP):
            pv_range  = pivot_sheet.range(pos).expand()
            pv_values = pv_range.value

            if idx == 0:
                out_wb.sheets[0].name = tab_name
                tgt = out_wb.sheets[0]
            else:
                tgt = out_wb.sheets.add(name=tab_name, after=out_wb.sheets[-1])

            tgt.range("A1").value = pv_values

            if idx == NII_PIVOT_IDX:
                nii_data = pv_values
            elif idx == FEE_PIVOT_IDX:
                fee_data = pv_values

        out_wb.save(out_path)
        out_wb.close()
        out_app.quit()
        log.info(f"Output file saved: {out_path}")

        # Save QS file
        wb.save()
        wb.close()
        log.info("QS SQL download.xlsx saved ✓")

    finally:
        app.quit()

    # ── Step 7: MoM analysis & email ──────────────────────────────────────────
    nii_text = get_mom_contributors(nii_data, "NII", yyyymm)
    fee_text = get_mom_contributors(fee_data, "Fee", yyyymm)

    log.info("MoM analysis complete, sending email …")
    send_email(out_path, yyyymm, nii_text, fee_text)

    log.info(f"── Pipeline complete for {yyyymm} ✓ ──")


# ══════════════════════════════════════════════════════════════════════════════
# FILE WATCHER
# ══════════════════════════════════════════════════════════════════════════════

class DropHandler(FileSystemEventHandler):
    """Trigger pipeline when a YYYYMM.xlsx file is created in the watch folder."""

    def on_created(self, event):
        if event.is_directory:
            return

        fname = os.path.basename(event.src_path)
        match = re.match(r"^(\d{6})\.xlsx$", fname, re.IGNORECASE)
        if not match:
            return

        yyyymm = int(match.group(1))
        log.info(f"Detected drop: {fname}  (period {yyyymm})")
        time.sleep(3)   # wait for file copy to complete before opening

        try:
            process(event.src_path, yyyymm)
        except Exception as exc:
            log.exception(f"Pipeline failed for {yyyymm}: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    log.info(f"PF Automation started.  Watching: {WATCH_FOLDER}")
    log.info("Drop a YYYYMM.xlsx file into the folder to trigger the pipeline.")
    log.info("Press Ctrl+C to stop.\n")

    handler  = DropHandler()
    observer = PollingObserver()          # PollingObserver handles UNC/network paths
    observer.schedule(handler, WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("Stopping watcher …")
        observer.stop()

    observer.join()
    log.info("Watcher stopped.")
