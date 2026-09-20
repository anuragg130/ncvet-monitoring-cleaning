"""Read-only presentation layer for the existing monitoring workbooks."""
from pathlib import Path
import json
import pandas as pd
import openpyxl

ROOT = Path(__file__).resolve().parent
RUNS = ROOT / "Runs"


def read_table(path, sheet, header_marker=None):
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in workbook.sheetnames:
            return pd.DataFrame()
        rows = list(workbook[sheet].values)
    finally:
        workbook.close()
    index = 0
    if header_marker:
        index = next((i for i, row in enumerate(rows) if header_marker in row), -1)
        if index < 0:
            raise ValueError(f"Could not locate the {sheet} table header.")
    if not rows:
        return pd.DataFrame()
    headers = [str(v) if v is not None else f"Column {i+1}" for i, v in enumerate(rows[index])]
    return pd.DataFrame([r for r in rows[index+1:] if any(v is not None for v in r)], columns=headers).fillna("")


def available_runs():
    return sorted([p for p in RUNS.glob("run_*") if
                   (p / "AA_Cleaned_Master.xlsx").is_file() and
                   (p / "AA_Consolidated_Master.xlsx").is_file()], reverse=True)


def load_run(folder):
    folder = Path(folder)
    raw = read_table(folder / "AA_Consolidated_Master.xlsx", "Master Data")
    clean = read_table(folder / "AA_Cleaned_Master.xlsx", "Master Data")
    if len(raw) != len(clean):
        raise ValueError("Raw and cleaned row counts differ. This run cannot be compared by row.")
    for column in ["Source File", "Source Sheet", "Source Row"]:
        if column in raw and column in clean and not raw[column].equals(clean[column]):
            raise ValueError("Source row alignment differs between the workbooks.")
    raw.index = pd.Index(range(2, len(raw)+2), name="Workbook row")
    clean.index = raw.index
    log = read_table(folder / "AA_Cleaned_Master.xlsx", "Cleaning Log", "Master Data Row")
    log["Master Data Row"] = pd.to_numeric(log["Master Data Row"], errors="raise").astype(int)
    agency_column = "Assessment Agency (Master)" if "Assessment Agency (Master)" in raw else "Assessment Agency Name"
    log["Agency"] = log["Master Data Row"].map(raw[agency_column]).fillna("Unknown")
    status_path = folder / "AA_Status_Report.xlsx"
    status = read_table(status_path, "Validation Details", "S.No.") if status_path.exists() else pd.DataFrame()
    responses = read_table(status_path, "Response Log", "Response Row") if status_path.exists() else pd.DataFrame()
    manifest_path = folder / "pipeline_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
    period = manifest.get("reporting_month_setting") or manifest.get("reporting_month")
    if not period and "Reporting Month" in raw and len(raw):
        value = raw["Reporting Month"].iloc[0]
        parsed = pd.to_datetime(value, errors="coerce")
        period = parsed.strftime("%B %Y") if not pd.isna(parsed) else str(value)
    return dict(raw=raw, clean=clean, log=log, status=status, responses=responses,
                manifest=manifest, period=period or "Selected reporting period", agency_column=agency_column)


def metrics(raw, clean, log):
    unresolved = log["Match Method"].eq("Unmatched")
    changed = log["Original Value"].astype(str).ne(log["Cleaned Value"].astype(str))
    return dict(rows=len(raw), changes=int(changed.sum()),
                resolved=int((changed & ~unresolved).sum()),
                missing=int((unresolved & log["Cleaned Value"].eq("Missing")).sum()),
                unmatched=int(unresolved.sum()),
                flagged=int(clean["Cleaning Flags"].astype(str).str.strip().ne("").sum()),
                affected=int(log.loc[changed, "Master Data Row"].nunique()),
                fuzzy=int(log["Match Method"].eq("Fuzzy").sum()))
