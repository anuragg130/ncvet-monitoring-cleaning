"""NCVET submission validation and cleaning pipeline. Run python run_pipeline.py --links FILE --month MONTH."""

from __future__ import annotations

# =============================================================================
# USER SETTINGS - edit these before running
# =============================================================================
from pathlib import Path
import os

# Everything lives here by default. Change this one line if the folder moves.
MONITORING_FOLDER = Path(__file__).resolve().parent

# ---- This month's input ------------------------------------------------------
# The Google Form / Sheet export with one row per AA submission. It needs an
# agency column, a month column, a "have assessments been conducted" column and
# a link column. A Timestamp column is used when present but is not required.
LINK_SHEET = MONITORING_FOLDER / "Links-July.xlsx"

# Blank = use the latest month found in the data. Otherwise "July 2026",
# "Jul-26" or "07/2026". Setting it also keeps a zero-response month visible.
REPORTING_MONTH = ""

# ---- Master references (rarely change) ---------------------------------------
DCF_TEMPLATE = MONITORING_FOLDER / "AA Monthly DCF 2627 - V1 - 06 May 2026  (1).xlsx"
AA_LIST = MONITORING_FOLDER / "AA List.xlsx"
SECTORS_LIST = MONITORING_FOLDER / "Sectors_List.xlsx"
STATE_DISTRICT_LIST = MONITORING_FOLDER / "State District List.xlsx"

# ---- Output ------------------------------------------------------------------
OUTPUT_ROOT = MONITORING_FOLDER / "Runs"

# ---- Google Drive access for private Form uploads ----------------------------
# Public / link-accessible files need neither. For private uploads set one:
GOOGLE_DRIVE_ACCESS_TOKEN = os.environ.get("GOOGLE_DRIVE_ACCESS_TOKEN", "")
GOOGLE_SERVICE_ACCOUNT_JSON = Path(r"")

# Corporate proxies commonly break certificate chains. Kept False to match the
# existing scripts; set True on a network with a clean chain.
VERIFY_SSL_CERTIFICATES = os.environ.get("NCVET_WEB_MODE") == "1"

# ---- Behaviour ---------------------------------------------------------------
# Used only when the month cell says just "July". None = take the year from the
# response Timestamp.
DEFAULT_REPORTING_YEAR: int | None = None

# Stages to run: 1 download/combine/status, 2 clean.
STAGES = (1, 2)

# Set this to an existing AA_Consolidated_Master.xlsx to re-run cleaning and
# without downloading anything again. Implies stage 2.
EXISTING_COMBINED_WORKBOOK = Path(r"")

# Stop the whole run if a stage fails, rather than continuing with stale inputs.
STOP_ON_STAGE_FAILURE = True
# =============================================================================

import argparse
import importlib
import json
import logging
import sys
import traceback
from datetime import datetime
from typing import Any, Sequence

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

APP_NAME = "NCVET AA Monthly Pipeline"
APP_VERSION = "1.0.0"

STAGE_NAMES = {
    1: "Download, validate, combine and status report",
    2: "Clean against master lists and dates",
}


def log_banner(text: str) -> None:
    line = "-" * max(12, min(78, len(text) + 4))
    print(f"\n{line}\n  {text}\n{line}")


def require_file(path: Path, label: str) -> Path:
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"{label} does not point to a file: {path}")
    return path


def safe_month_tag(value: str) -> str:
    text = "".join(ch if ch.isalnum() or ch in " -_" else "-" for ch in str(value)).strip()
    return " ".join(text.split()).replace(" ", "_") or "Month"


# -----------------------------------------------------------------------------
# Stage 1 — download, validate, combine, status report
# -----------------------------------------------------------------------------
def stage_download(run_folder: Path) -> dict[str, Any]:
    aa_monitor = importlib.import_module("aa_monitor")

    aa_monitor.CREATE_TIMESTAMPED_RUN_FOLDER = False
    aa_monitor.VERIFY_SSL_CERTIFICATES = VERIFY_SSL_CERTIFICATES
    aa_monitor.GOOGLE_DRIVE_ACCESS_TOKEN = GOOGLE_DRIVE_ACCESS_TOKEN
    aa_monitor.GOOGLE_SERVICE_ACCOUNT_JSON = Path(GOOGLE_SERVICE_ACCOUNT_JSON)
    aa_monitor.DEFAULT_REPORTING_YEAR = DEFAULT_REPORTING_YEAR
    aa_monitor.REPORTING_MONTHS = [REPORTING_MONTH] if REPORTING_MONTH else []

    status_report, combined_master = aa_monitor.run(
        responses_path=require_file(LINK_SHEET, "LINK_SHEET"),
        dcf_path=require_file(DCF_TEMPLATE, "DCF_TEMPLATE"),
        master_path=require_file(AA_LIST, "AA_LIST"),
        output_root=run_folder,
    )
    return {
        "status_report": Path(status_report),
        "combined_master": Path(combined_master),
    }


# -----------------------------------------------------------------------------
# Stage 2 — clean
# -----------------------------------------------------------------------------
def stage_clean(combined_master: Path, run_folder: Path) -> dict[str, Any]:
    cleaner = importlib.import_module("clean_aa_data")

    cleaner.AA_LIST_PATH = require_file(AA_LIST, "AA_LIST")
    cleaner.SECTORS_LIST_PATH = require_file(SECTORS_LIST, "SECTORS_LIST")
    cleaner.STATE_DISTRICT_PATH = require_file(STATE_DISTRICT_LIST, "STATE_DISTRICT_LIST")
    cleaner.DCF_TEMPLATE_PATH = require_file(DCF_TEMPLATE, "DCF_TEMPLATE")
    cleaner.REPORTING_MONTH_HINT = REPORTING_MONTH

    output = run_folder / "AA_Cleaned_Master.xlsx"
    summary = cleaner.clean_workbook(
        input_path=require_file(combined_master, "combined master"),
        output_path=output,
        preferred_sheet="",
    )
    summary["cleaned_master"] = output
    return summary


# -----------------------------------------------------------------------------
# Runner
# -----------------------------------------------------------------------------
def run_pipeline(stages: Sequence[int], from_combined: Path | None) -> int:
    started = datetime.now()
    run_folder = OUTPUT_ROOT / f"run_{started.strftime('%Y%m%d_%H%M%S')}"
    run_folder.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(run_folder / "pipeline.log", encoding="utf-8"),
        ],
        force=True,
    )

    print(f"{APP_NAME} v{APP_VERSION}")
    print(f"Run folder: {run_folder}")
    print(f"Stages: {', '.join(str(stage) for stage in stages)}")
    if REPORTING_MONTH:
        print(f"Reporting month: {REPORTING_MONTH}")

    manifest: dict[str, Any] = {
        "app": APP_NAME,
        "version": APP_VERSION,
        "started_at": started.isoformat(timespec="seconds"),
        "run_folder": str(run_folder),
        "link_sheet": str(LINK_SHEET),
        "reporting_month_setting": REPORTING_MONTH or "(latest found)",
        "masters": {
            "dcf_template": str(DCF_TEMPLATE),
            "aa_list": str(AA_LIST),
            "sectors_list": str(SECTORS_LIST),
            "state_district_list": str(STATE_DISTRICT_LIST),
        },
        "stages": {},
    }

    status_report: Path | None = None
    combined_master: Path | None = Path(from_combined) if from_combined else None
    cleaned_master: Path | None = None
    failures = 0

    for stage in stages:
        log_banner(f"Stage {stage} — {STAGE_NAMES[stage]}")
        stage_started = datetime.now()
        try:
            if stage == 1:
                result = stage_download(run_folder)
                status_report = result["status_report"]
                combined_master = result["combined_master"]
                print(f"  Status report    : {status_report}")
                print(f"  Combined master  : {combined_master}")
                manifest["stages"]["1_download_combine_status"] = {
                    "status_report": str(status_report),
                    "combined_master": str(combined_master),
                }

            elif stage == 2:
                if not combined_master:
                    raise FileNotFoundError(
                        "No combined master available. Run stage 1, or set "
                        "EXISTING_COMBINED_WORKBOOK / --from-combined."
                    )
                summary = stage_clean(combined_master, run_folder)
                cleaned_master = summary["cleaned_master"]
                print(f"  Rows cleaned     : {summary['rows']:,}")
                print(f"  Changes applied  : {summary['changes']:,}")
                print(f"  Rows flagged     : {summary['rows_with_flags']:,}")
                print(f"  Unmatched values : {summary['unmatched_values']:,}")
                print(f"  Cleaned master   : {cleaned_master}")
                manifest["stages"]["2_clean"] = {
                    key: (str(value) if isinstance(value, Path) else value)
                    for key, value in summary.items()
                }

            manifest["stages"].setdefault(f"stage_{stage}", {})
            elapsed = (datetime.now() - stage_started).total_seconds()
            print(f"  Stage {stage} completed in {elapsed:.1f}s")

        except Exception as exc:
            failures += 1
            logging.error("Stage %d failed: %s", stage, exc)
            traceback.print_exc()
            manifest["stages"][f"stage_{stage}_error"] = f"{type(exc).__name__}: {exc}"
            if STOP_ON_STAGE_FAILURE:
                break

    manifest["finished_at"] = datetime.now().isoformat(timespec="seconds")
    manifest["elapsed_seconds"] = round((datetime.now() - started).total_seconds(), 1)
    manifest["failed_stages"] = failures
    (run_folder / "pipeline_manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    log_banner("Pipeline finished" if not failures else "Pipeline finished with errors")
    print(f"Everything for this month is in: {run_folder}")
    if failures:
        print(f"{failures} stage(s) failed — see pipeline.log and the traceback above.")
    return 1 if failures else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the NCVET AA monthly monitoring pipeline.")
    parser.add_argument("--links", type=Path, default=None, help="This month's link sheet")
    parser.add_argument("--month", default=None, help="Reporting month, e.g. 'July 2026'")
    parser.add_argument("--output", type=Path, default=None, help="Root folder for run folders")
    parser.add_argument("--dcf", type=Path, default=None, help="AA Monthly DCF template")
    parser.add_argument("--aa-list", type=Path, default=None, help="Approved AA master list")
    parser.add_argument("--sectors", type=Path, default=None, help="Sectors master list")
    parser.add_argument("--state-district", type=Path, default=None, help="State/District master list")
    parser.add_argument("--token", default=None, help="Google Drive OAuth access token")
    parser.add_argument(
        "--stages", default=None,
        help="Comma-separated stages to run, e.g. '2'. Default 1,2.",
    )
    parser.add_argument(
        "--from-combined", type=Path, default=None,
        help="Skip downloading and clean this existing combined master.",
    )
    return parser


def apply_overrides(args: argparse.Namespace) -> tuple[list[int], Path | None]:
    global LINK_SHEET, REPORTING_MONTH, OUTPUT_ROOT, DCF_TEMPLATE, AA_LIST
    global SECTORS_LIST, STATE_DISTRICT_LIST, GOOGLE_DRIVE_ACCESS_TOKEN

    if args.links:
        LINK_SHEET = Path(args.links)
    if args.month is not None:
        REPORTING_MONTH = args.month
    if args.output:
        OUTPUT_ROOT = Path(args.output)
    if args.dcf:
        DCF_TEMPLATE = Path(args.dcf)
    if args.aa_list:
        AA_LIST = Path(args.aa_list)
    if args.sectors:
        SECTORS_LIST = Path(args.sectors)
    if args.state_district:
        STATE_DISTRICT_LIST = Path(args.state_district)
    if args.token:
        GOOGLE_DRIVE_ACCESS_TOKEN = args.token

    from_combined = args.from_combined or (
        EXISTING_COMBINED_WORKBOOK if str(EXISTING_COMBINED_WORKBOOK) not in ("", ".") else None
    )
    if from_combined and not Path(from_combined).is_file():
        raise FileNotFoundError(f"Combined workbook not found: {from_combined}")

    if args.stages:
        stages = [int(part) for part in args.stages.replace(" ", "").split(",") if part]
    elif from_combined:
        stages = [2]
    else:
        stages = list(STAGES)

    invalid = [stage for stage in stages if stage not in STAGE_NAMES]
    if invalid:
        raise ValueError(f"Unknown stage(s): {invalid}. Valid stages are 1 and 2.")
    return sorted(set(stages)), (Path(from_combined) if from_combined else None)


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        stages, from_combined = apply_overrides(args)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    return run_pipeline(stages, from_combined)


if __name__ == "__main__":
    raise SystemExit(main())
