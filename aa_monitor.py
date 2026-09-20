"""NCVET Assessment Agency submission downloader and format monitor.

Edit the USER SETTINGS block below, then run:
    python aa_monitor.py

Command-line overrides are also available; run `python aa_monitor.py --help`.
"""

from __future__ import annotations

# =============================================================================
# USER SETTINGS — edit these paths/options before running the script
# =============================================================================
from pathlib import Path

FORM_RESPONSES_PATH = Path(r"C:\PATH\TO\Google Forms responses.xlsx")
DCF_TEMPLATE_PATH = Path(r"C:\Users\Anurag\Downloads\AA DCF.xlsx")
AA_CATEGORIES_PATH = Path(
    r"C:\Users\Anurag\Documents\NCVET BD\Innovation\Dashboard Collaterals\aa cateogries.xlsx"
)
OUTPUT_ROOT = Path(
    r"C:\Users\Anurag\Documents\NCVET BD\Innovation\Dashboard Collaterals\AA Monitoring Output"
)

# Leave blank to infer months from the response file. To include a month even
# when nobody responded, enter values such as ["Apr-26", "May-26", "Jun-26"].
REPORTING_MONTHS: list[str] = []

# Used only when the Month cell says just "April", "May", etc. Leave as None
# to use the year from that response's Timestamp.
DEFAULT_REPORTING_YEAR: int | None = None

# A fresh timestamped run folder prevents older downloads from being overwritten.
CREATE_TIMESTAMPED_RUN_FOLDER = True

# The script creates this separate workbook by aligning every parseable AA data
# row to the DCF columns. Files marked Not in Format are included with blank
# cells for their missing columns, so usable data is not discarded.
CONSOLIDATED_MASTER_FILENAME = "AA_Consolidated_Master.xlsx"

# Optional authentication for private Google Drive uploads. Either paste a
# temporary OAuth access token or point to a service-account JSON file and share
# the Drive files/folder with that service account. Public links need neither.
GOOGLE_DRIVE_ACCESS_TOKEN = ""
GOOGLE_SERVICE_ACCOUNT_JSON = Path(r"")

# Matching/validation controls.
FUZZY_AGENCY_MATCH_THRESHOLD = 0.90
FUZZY_AGENCY_MATCH_MARGIN = 0.03
DCF_SHEET_NAME = "AA Monitoring"
DCF_HEADER_ROW = 1
HEADER_SCAN_ROWS = 30
MAX_COLUMNS_TO_SCAN = 300
DOWNLOAD_TIMEOUT_SECONDS = 180

# SSL certificate verification is intentionally disabled for every network
# path in this script (session defaults, individual requests, Google token
# refresh, and Python's standard HTTPS context).
VERIFY_SSL_CERTIFICATES = False

# Expected form-response headers.
TIMESTAMP_HEADER = "Timestamp"
AGENCY_HEADER = "Name of Assessment Agency"
MONTH_HEADER = "Month"
ASSESSMENT_CONDUCTED_HEADER = "Have Assessments been conducted for the Month"
LINK_HEADER = "Link"

# Accepted spellings for each form-response column. The first alias that is
# present wins; add new wordings here when the Google Form changes.
RESPONSE_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "timestamp": ("Timestamp", "Submission Time", "Date of Submission"),
    "agency": (
        "Name of Assessment Agency",
        "Name of the Assessment Agency",
        "Assessment Agency Name",
        "Assessment Agency",
    ),
    "month": (
        "Month",
        "Month for which data is submitted",
        "Month for which the data is submitted",
        "Reporting Month",
        "Month of Batch Assessment",
    ),
    "conducted": (
        "Have Assessments been conducted for the Month",
        "Have Assessments been conducted for the month",
        "Have Assessments been conducted",
        "Assessments Conducted",
    ),
    "link": ("Link", "File Link", "Data Link", "Upload", "Upload the file"),
}

# Fields the script can work without. A missing Timestamp means the last row
# for an agency/month wins instead of the latest submission time.
OPTIONAL_RESPONSE_FIELDS = {"timestamp"}

# Agency names the form and the AA master genuinely spell differently. The key
# is the name as submitted on the form, the value is the name exactly as it
# appears in AA List.xlsx. Punctuation and case do not matter.
AGENCY_NAME_OVERRIDES: dict[str, str] = {
    # "Some Agency As Typed On The Form": "Some Agency Pvt. Ltd.",
}
# =============================================================================

import argparse
import csv
import html
import json
import logging
import os
import re
import shutil
import ssl
import sys
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from difflib import SequenceMatcher
from email.message import Message
from html.parser import HTMLParser
from typing import Any, Iterable, Iterator
from urllib.parse import parse_qs, unquote, urljoin, urlparse

try:
    import requests
except ImportError as exc:  # pragma: no cover - friendly startup failure
    raise SystemExit(
        "Missing dependency 'requests'. Run: pip install -r requirements.txt"
    ) from exc

if not VERIFY_SSL_CERTIFICATES:
    # Cover urllib/standard-library HTTPS users as well as requests, and avoid
    # repeated InsecureRequestWarning messages for the explicitly unverified
    # requests below.
    os.environ["PYTHONHTTPSVERIFY"] = "0"
    ssl._create_default_https_context = ssl._create_unverified_context
    requests.packages.urllib3.disable_warnings(
        requests.packages.urllib3.exceptions.InsecureRequestWarning
    )

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.workbook.properties import CalcProperties
except ImportError as exc:  # pragma: no cover - friendly startup failure
    raise SystemExit(
        "Missing dependency 'openpyxl'. Run: pip install -r requirements.txt"
    ) from exc


STATUS_SUBMITTED = "Submitted"
STATUS_NOT_IN_FORMAT = "Not in Format"
STATUS_NO_ASSESSMENTS = "No Assessments"
STATUS_DATA_NOT_SUBMITTED = "Data Not Submitted"
STATUS_FILE_ERROR = "File Error"
STATUS_ORDER = [
    STATUS_SUBMITTED,
    STATUS_NOT_IN_FORMAT,
    STATUS_NO_ASSESSMENTS,
    STATUS_DATA_NOT_SUBMITTED,
    STATUS_FILE_ERROR,
]
STATUS_COLORS = {
    STATUS_SUBMITTED: "C6E0B4",
    STATUS_NOT_IN_FORMAT: "FFD966",
    STATUS_NO_ASSESSMENTS: "B4C6E7",
    STATUS_DATA_NOT_SUBMITTED: "F4B6B6",
    STATUS_FILE_ERROR: "D9E1F2",
}

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
EXCEL_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
URL_RE = re.compile(r"https?://[^\s,;]+", re.IGNORECASE)
MONTH_NAMES = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


@dataclass(frozen=True, order=True)
class ReportingMonth:
    year: int
    month: int

    @property
    def label(self) -> str:
        return date(self.year, self.month, 1).strftime("%b-%y")

    @property
    def folder_name(self) -> str:
        return date(self.year, self.month, 1).strftime("%Y-%m_%b")


@dataclass
class Agency:
    serial_no: Any
    name: str
    jurisdiction: str
    normalized_name: str


@dataclass
class ResponseRecord:
    row_number: int
    timestamp_raw: Any
    timestamp: datetime | None
    agency_raw: str
    month_raw: Any
    month: ReportingMonth | None
    conducted_raw: str
    link_raw: str
    links: list[str]
    matched_agency: Agency | None = None
    match_method: str = ""
    match_score: float | None = None
    selected: bool = False
    response_note: str = ""


@dataclass
class FileValidation:
    file_path: Path
    sheet_name: str = ""
    header_row: int | None = None
    matched_column_count: int = 0
    missing_columns: list[str] = field(default_factory=list)
    error: str = ""
    rows_consolidated: int = 0
    consolidation_error: str = ""


@dataclass
class StatusEvidence:
    status: str
    response: ResponseRecord | None = None
    downloaded_files: list[Path] = field(default_factory=list)
    validations: list[FileValidation] = field(default_factory=list)
    missing_columns: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


class DownloadConfirmationParser(HTMLParser):
    """Extract the first form and hidden inputs from a Drive warning page."""

    def __init__(self) -> None:
        super().__init__()
        self.action = ""
        self.inputs: dict[str, str] = {}
        self._inside_form = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = dict(attrs)
        if tag.lower() == "form" and not self.action:
            self.action = html.unescape(values.get("action") or "")
            self._inside_form = True
        elif tag.lower() == "input" and self._inside_form:
            name = values.get("name")
            if name:
                self.inputs[name] = html.unescape(values.get("value") or "")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "form" and self._inside_form:
            self._inside_form = False


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).replace("\xa0", " ")
    return " ".join(text.split()).strip()


def normalize_header(value: Any) -> str:
    return "".join(ch for ch in normalize_text(value).casefold() if ch.isalnum())


# Word forms that mean the same thing in an Assessment Agency name. Both the
# form response and the master list are folded through this map before matching,
# so "Private Limited" and "Pvt. Ltd." land on the same key.
AGENCY_WORD_MAP = {
    "private": "pvt", "limited": "ltd", "technologies": "technology",
    "services": "service", "solutions": "solution", "consultants": "consultant",
    "consultancy": "consultant", "skills": "skill", "skilling": "skill",
    "assessments": "assessment", "assessors": "assessor", "assessor": "assessor",
    "centers": "centre", "center": "centre", "centres": "centre",
    "organisation": "organization", "corporation": "corp", "company": "co",
    "incorporated": "inc", "exporters": "export", "exports": "export",
    "exporter": "export", "garments": "garment", "manufacturers": "mfg",
    "manufacturer": "mfg", "manufacturing": "mfg", "institutes": "institute",
    "foundations": "foundation", "societies": "society",
}

# Dropped before matching: they never distinguish one agency from another.
AGENCY_STOPWORDS = {"the", "of", "and", "a", "an", "for", "consortium"}


def normalize_agency_name(value: Any) -> str:
    """Fold an agency name to a comparison key.

    Parenthetical qualifiers such as "(Consortium)" or "(GEMA)" are dropped,
    common word variants are folded, and punctuation and case are ignored.
    """
    text = normalize_text(value)
    if not text:
        return ""
    override = AGENCY_NAME_OVERRIDES.get(text) or AGENCY_NAME_OVERRIDES.get(text.casefold())
    if override:
        text = normalize_text(override)
    text = re.sub(r"\(.*?\)", " ", text)
    text = "".join(ch if ch.isalnum() else " " for ch in text.casefold())
    words = [AGENCY_WORD_MAP.get(word, word) for word in text.split()]
    return "".join(word for word in words if word not in AGENCY_STOPWORDS)


def safe_filename(value: str, fallback: str = "file") -> str:
    cleaned = INVALID_FILENAME_CHARS.sub("_", normalize_text(value)).strip(" ._")
    cleaned = re.sub(r"_+", "_", cleaned)
    return (cleaned[:180] or fallback).strip(" .")


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for number in range(2, 10000):
        candidate = path.with_name(f"{path.stem}_{number}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Could not create a unique file name for {path}")


def parse_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    text = normalize_text(value)
    if not text:
        return None
    candidates = [text, text.replace("Z", "+00:00")]
    for candidate in candidates:
        try:
            return datetime.fromisoformat(candidate)
        except ValueError:
            pass
    for pattern in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            continue
    return None


def parse_reporting_month(
    value: Any,
    timestamp: datetime | None = None,
    default_year: int | None = None,
) -> ReportingMonth | None:
    if isinstance(value, (datetime, date)):
        return ReportingMonth(value.year, value.month)
    if isinstance(value, (int, float)) and 1 <= int(value) <= 12:
        year = default_year or (timestamp.year if timestamp else None)
        return ReportingMonth(year, int(value)) if year else None

    text = normalize_text(value).casefold()
    if not text:
        return None
    text = text.replace(",", " ").replace("_", " ")
    text = re.sub(r"\s+", " ", text).strip()

    numeric = re.fullmatch(r"(20\d{2})[-/]([01]?\d)", text)
    if numeric and 1 <= int(numeric.group(2)) <= 12:
        return ReportingMonth(int(numeric.group(1)), int(numeric.group(2)))
    numeric = re.fullmatch(r"([01]?\d)[-/](20\d{2})", text)
    if numeric and 1 <= int(numeric.group(1)) <= 12:
        return ReportingMonth(int(numeric.group(2)), int(numeric.group(1)))

    match = re.fullmatch(r"([a-z]+)[\s\-/]*(\d{2,4})?", text)
    if match:
        month_no = MONTH_NAMES.get(match.group(1))
        if month_no:
            year_text = match.group(2)
            if year_text:
                year = int(year_text)
                if year < 100:
                    year += 2000
            else:
                year = default_year or (timestamp.year if timestamp else None)
            return ReportingMonth(year, month_no) if year else None

    match = re.fullmatch(r"(\d{2,4})[\s\-/]*([a-z]+)", text)
    if match and match.group(2) in MONTH_NAMES:
        year = int(match.group(1))
        if year < 100:
            year += 2000
        return ReportingMonth(year, MONTH_NAMES[match.group(2)])
    return None


def extract_links(value: Any, hyperlink_target: str = "") -> list[str]:
    text = normalize_text(value)
    candidates: list[str] = []
    if hyperlink_target:
        candidates.append(hyperlink_target.strip())

    hyperlink_formula = re.match(
        r'^=HYPERLINK\(\s*"([^"]+)"', text, flags=re.IGNORECASE
    )
    if hyperlink_formula:
        candidates.append(hyperlink_formula.group(1))

    candidates.extend(URL_RE.findall(text))
    if not candidates and text:
        for piece in re.split(r"[\n,;]+", text):
            piece = piece.strip()
            if piece and (Path(piece).exists() or piece.lower().startswith("file:")):
                candidates.append(piece)

    result: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        cleaned = candidate.strip().rstrip(").]}>\"'")
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)
    return result


def header_map(headers: Iterable[Any]) -> dict[str, int]:
    result: dict[str, int] = {}
    for index, header in enumerate(headers):
        normalized = normalize_header(header)
        if normalized and normalized not in result:
            result[normalized] = index
    return result


def _augment_header_mapping(mapping: dict[str, int], aliases: set[str]) -> dict[str, int]:
    """Tolerate decorated headers such as 'Assessment Agency Name (68)'."""
    result = dict(mapping)
    for alias in aliases:
        if alias in result or not alias:
            continue
        for header, index in mapping.items():
            if not header:
                continue
            if header.startswith(alias) or (len(header) >= 4 and alias.startswith(header)):
                result[alias] = index
                break
    return result


def load_agency_master(path: Path) -> list[Agency]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        rows = list(sheet.iter_rows(min_row=1, max_row=10, values_only=True))
        name_aliases = {normalize_header(x) for x in ("Name", "Assessment Agency Name")}
        jurisdiction_aliases = {normalize_header(x) for x in ("Jurisdiction", "Category")}
        serial_aliases = {normalize_header(x) for x in ("Sr. No.", "S.No.", "Serial No")}
        located: tuple[int, dict[str, int]] | None = None
        for row_number, row in enumerate(rows, start=1):
            mapping = _augment_header_mapping(
                header_map(row), name_aliases | jurisdiction_aliases | serial_aliases
            )
            if any(alias in mapping for alias in name_aliases):
                located = (row_number, mapping)
                break
        if not located:
            raise ValueError("Could not find a Name/Assessment Agency Name column in the master list.")
        header_row, mapping = located
        name_index = next(mapping[x] for x in name_aliases if x in mapping)
        jurisdiction_index = next(
            (mapping[x] for x in jurisdiction_aliases if x in mapping), None
        )
        serial_index = next((mapping[x] for x in serial_aliases if x in mapping), None)

        agencies: list[Agency] = []
        seen: set[str] = set()
        for excel_row, row in enumerate(
            sheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1
        ):
            name = normalize_text(row[name_index] if name_index < len(row) else "")
            if not name:
                continue
            normalized = normalize_agency_name(name)
            if normalized in seen:
                raise ValueError(f"Duplicate agency in master list after normalization: {name}")
            seen.add(normalized)
            serial = (
                row[serial_index]
                if serial_index is not None and serial_index < len(row)
                else len(agencies) + 1
            )
            jurisdiction = (
                normalize_text(row[jurisdiction_index])
                if jurisdiction_index is not None and jurisdiction_index < len(row)
                else ""
            )
            agencies.append(Agency(serial, name, jurisdiction, normalized))
        if not agencies:
            raise ValueError("The agency master contains no agency rows.")
        return agencies
    finally:
        workbook.close()


def load_required_columns(path: Path) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if DCF_SHEET_NAME and DCF_SHEET_NAME in workbook.sheetnames:
            sheets = [workbook[DCF_SHEET_NAME]]
        else:
            sheets = [workbook[name] for name in workbook.sheetnames]

        best_headers: list[str] = []
        for sheet in sheets:
            if DCF_HEADER_ROW:
                rows = sheet.iter_rows(
                    min_row=DCF_HEADER_ROW,
                    max_row=DCF_HEADER_ROW,
                    max_col=min(sheet.max_column, MAX_COLUMNS_TO_SCAN),
                    values_only=True,
                )
            else:
                rows = sheet.iter_rows(
                    min_row=1,
                    max_row=min(HEADER_SCAN_ROWS, sheet.max_row),
                    max_col=min(sheet.max_column, MAX_COLUMNS_TO_SCAN),
                    values_only=True,
                )
            for row in rows:
                headers = [normalize_text(value) for value in row if normalize_text(value)]
                if len(headers) > len(best_headers):
                    best_headers = headers
        if not best_headers:
            raise ValueError("No required headers were found in the DCF template.")
        normalized = [normalize_header(value) for value in best_headers]
        duplicates = sorted({x for x in normalized if normalized.count(x) > 1})
        if duplicates:
            raise ValueError("The DCF template contains duplicate required headers.")
        return best_headers
    finally:
        workbook.close()


def resolve_response_columns(headers) -> dict[str, int]:
    """Map each logical response field to a column index, using the aliases."""
    mapping = header_map(headers)
    resolved: dict[str, int] = {}
    for field, aliases in RESPONSE_HEADER_ALIASES.items():
        index = None
        for alias in aliases:
            key = normalize_header(alias)
            if key in mapping:
                index = mapping[key]
                break
        if index is None:
            for alias in aliases:
                key = normalize_header(alias)
                if not key:
                    continue
                for header_key, header_index in mapping.items():
                    if len(header_key) < 5:
                        continue
                    if header_key.startswith(key) or key.startswith(header_key):
                        index = header_index
                        break
                if index is not None:
                    break
        if index is not None and index not in resolved.values():
            resolved[field] = index
        elif index is not None:
            resolved[field] = index
    missing = [
        field
        for field in RESPONSE_HEADER_ALIASES
        if field not in resolved and field not in OPTIONAL_RESPONSE_FIELDS
    ]
    if missing:
        details = "; ".join(
            f"{field} (expected one of: {', '.join(RESPONSE_HEADER_ALIASES[field])})"
            for field in missing
        )
        raise ValueError(f"Missing form-response columns: {details}")
    for field in OPTIONAL_RESPONSE_FIELDS:
        if field not in resolved:
            logging.warning(
                "Form responses have no '%s' column; using sheet order to pick the "
                "latest submission for each agency/month.",
                field,
            )
    return resolved


def read_form_responses(path: Path) -> list[ResponseRecord]:
    records: list[ResponseRecord] = []

    if path.suffix.casefold() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        if not rows:
            raise ValueError("The form-response CSV is empty.")
        columns = resolve_response_columns(rows[0])

        for row_number, row in enumerate(rows[1:], start=2):
            def get_value(field: str) -> Any:
                index = columns.get(field)
                if index is None or index >= len(row):
                    return ""
                return row[index]

            timestamp_raw = get_value("timestamp")
            timestamp = parse_datetime(timestamp_raw)
            month_raw = get_value("month")
            link_raw = normalize_text(get_value("link"))
            agency_raw = normalize_text(get_value("agency"))
            conducted_raw = normalize_text(get_value("conducted"))
            if not any((timestamp_raw, agency_raw, month_raw, conducted_raw, link_raw)):
                continue
            records.append(
                ResponseRecord(
                    row_number=row_number,
                    timestamp_raw=timestamp_raw,
                    timestamp=timestamp,
                    agency_raw=agency_raw,
                    month_raw=month_raw,
                    month=parse_reporting_month(month_raw, timestamp, DEFAULT_REPORTING_YEAR),
                    conducted_raw=conducted_raw,
                    link_raw=link_raw,
                    links=extract_links(link_raw),
                )
            )
        return records

    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        header_cells = next(sheet.iter_rows(min_row=1, max_row=1))
        columns = resolve_response_columns(cell.value for cell in header_cells)
        indexes = {field: index + 1 for field, index in columns.items()}

        def cell_for(row_number: int, field: str):
            column = indexes.get(field)
            return sheet.cell(row_number, column) if column else None

        for row_number in range(2, sheet.max_row + 1):
            timestamp_cell = cell_for(row_number, "timestamp")
            timestamp_raw = timestamp_cell.value if timestamp_cell else None
            agency_cell = cell_for(row_number, "agency")
            agency_raw = normalize_text(agency_cell.value if agency_cell else None)
            month_cell = cell_for(row_number, "month")
            month_raw = month_cell.value if month_cell else None
            conducted_cell = cell_for(row_number, "conducted")
            conducted_raw = normalize_text(conducted_cell.value if conducted_cell else None)
            link_cell = cell_for(row_number, "link")
            link_raw = normalize_text(link_cell.value if link_cell else None)
            hyperlink_target = (
                link_cell.hyperlink.target if link_cell is not None and link_cell.hyperlink else ""
            )
            if not any((timestamp_raw, agency_raw, month_raw, conducted_raw, link_raw, hyperlink_target)):
                continue
            timestamp = parse_datetime(timestamp_raw)
            records.append(
                ResponseRecord(
                    row_number=row_number,
                    timestamp_raw=timestamp_raw,
                    timestamp=timestamp,
                    agency_raw=agency_raw,
                    month_raw=month_raw,
                    month=parse_reporting_month(month_raw, timestamp, DEFAULT_REPORTING_YEAR),
                    conducted_raw=conducted_raw,
                    link_raw=link_raw or hyperlink_target,
                    links=extract_links(link_raw, hyperlink_target),
                )
            )
        return records
    finally:
        workbook.close()


def match_agencies(records: list[ResponseRecord], agencies: list[Agency]) -> None:
    exact = {agency.normalized_name: agency for agency in agencies}
    normalized_names = list(exact)
    for record in records:
        key = normalize_agency_name(record.agency_raw)
        if not key:
            record.response_note = "Agency name is blank"
            continue
        if key in exact:
            record.matched_agency = exact[key]
            record.match_method = "Exact"
            record.match_score = 1.0
            continue

        scored = sorted(
            ((SequenceMatcher(None, key, candidate).ratio(), candidate) for candidate in normalized_names),
            reverse=True,
        )
        best_score, best_key = scored[0]
        second_score = scored[1][0] if len(scored) > 1 else 0.0
        if (
            best_score >= FUZZY_AGENCY_MATCH_THRESHOLD
            and best_score - second_score >= FUZZY_AGENCY_MATCH_MARGIN
        ):
            record.matched_agency = exact[best_key]
            record.match_method = "Fuzzy"
            record.match_score = best_score
        else:
            record.match_method = "Unmatched"
            record.match_score = best_score
            record.response_note = "Agency name did not match the master list confidently"


def response_sort_key(record: ResponseRecord) -> tuple[datetime, int]:
    return (record.timestamp or datetime.min, record.row_number)


def select_latest_responses(
    records: list[ResponseRecord], reporting_months: list[ReportingMonth]
) -> dict[tuple[str, ReportingMonth], ResponseRecord]:
    permitted = set(reporting_months)
    grouped: dict[tuple[str, ReportingMonth], list[ResponseRecord]] = {}
    for record in records:
        if not record.month:
            record.response_note = (record.response_note + "; " if record.response_note else "") + "Invalid month"
            continue
        if record.month not in permitted:
            record.response_note = (
                (record.response_note + "; " if record.response_note else "")
                + "Outside configured reporting months"
            )
            continue
        if not record.matched_agency:
            continue
        key = (record.matched_agency.normalized_name, record.month)
        grouped.setdefault(key, []).append(record)

    selected: dict[tuple[str, ReportingMonth], ResponseRecord] = {}
    for key, group in grouped.items():
        latest = max(group, key=response_sort_key)
        latest.selected = True
        selected[key] = latest
        for record in group:
            if record is not latest:
                record.response_note = "Superseded by a later response for the same agency/month"
    return selected


def is_no_assessment(value: Any) -> bool:
    normalized = normalize_header(value)
    return normalized in {"no", "n", "false", "0", "noassessments", "noassessment"} or (
        normalized.startswith("no") and "assessment" in normalized
    )


def extract_google_file_id(url: str) -> str:
    parsed = urlparse(url)
    query_id = parse_qs(parsed.query).get("id", [""])[0]
    if query_id:
        return query_id
    patterns = (
        r"/file/d/([A-Za-z0-9_-]+)",
        r"/spreadsheets/d/([A-Za-z0-9_-]+)",
        r"/document/d/([A-Za-z0-9_-]+)",
        r"/presentation/d/([A-Za-z0-9_-]+)",
    )
    for pattern in patterns:
        match = re.search(pattern, parsed.path)
        if match:
            return match.group(1)
    return ""


def auth_headers() -> dict[str, str]:
    token = GOOGLE_DRIVE_ACCESS_TOKEN.strip() or os.getenv("GOOGLE_DRIVE_ACCESS_TOKEN", "").strip()
    if token:
        return {"Authorization": f"Bearer {token}"}
    credentials_path = str(GOOGLE_SERVICE_ACCOUNT_JSON).strip()
    if credentials_path and Path(credentials_path).is_file():
        try:
            from google.auth.transport.requests import Request
            from google.oauth2 import service_account
        except ImportError as exc:
            raise RuntimeError(
                "Google service-account authentication requires google-auth. "
                "Run: pip install -r requirements.txt"
            ) from exc
        credentials = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        refresh_session = requests.Session()
        refresh_session.verify = VERIFY_SSL_CERTIFICATES
        credentials.refresh(Request(session=refresh_session))
        return {"Authorization": f"Bearer {credentials.token}"}
    return {}


def filename_from_response(response: requests.Response, fallback: str) -> str:
    disposition = response.headers.get("Content-Disposition", "")
    if disposition:
        message = Message()
        message["content-disposition"] = disposition
        filename = message.get_filename()
        if filename:
            return safe_filename(unquote(filename), fallback)
    final_path = unquote(urlparse(response.url).path)
    candidate = Path(final_path).name
    return safe_filename(candidate, fallback) if candidate else fallback


def sniff_excel_extension(path: Path) -> str:
    with path.open("rb") as handle:
        prefix = handle.read(16)
    if prefix.startswith(b"PK") and zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            names = {name.casefold() for name in archive.namelist()}
        if "xl/workbook.bin" in names:
            return ".xlsb"
        if "xl/workbook.xml" in names:
            return ".xlsx"
    if prefix.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        return ".xls"
    if prefix.lstrip().startswith((b"<html", b"<!doctype html")):
        return ".html"
    return path.suffix.casefold()


def finalize_download(temp_path: Path, desired_path: Path) -> Path:
    detected = sniff_excel_extension(temp_path)
    if detected == ".html":
        preview = temp_path.read_text(encoding="utf-8", errors="ignore")[:300]
        temp_path.unlink(missing_ok=True)
        raise RuntimeError(f"The link returned an HTML/login page instead of a file: {preview!r}")
    if detected in {".xlsx", ".xlsb", ".xls"}:
        desired_path = desired_path.with_suffix(detected)
    desired_path = unique_path(desired_path)
    temp_path.replace(desired_path)
    return desired_path


def stream_response_to_file(response: requests.Response, desired_path: Path) -> Path:
    response.raise_for_status()
    temp_path = unique_path(desired_path.with_suffix(desired_path.suffix + ".part"))
    try:
        with temp_path.open("wb") as handle:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    handle.write(chunk)
        if temp_path.stat().st_size == 0:
            raise RuntimeError("The downloaded file is empty.")
        return finalize_download(temp_path, desired_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def download_google_drive(
    session: requests.Session, link: str, destination_dir: Path, index: int
) -> Path:
    file_id = extract_google_file_id(link)
    if not file_id:
        raise RuntimeError("Could not identify the Google Drive file ID.")
    headers = auth_headers()
    fallback = f"drive_file_{index}"

    if headers:
        metadata_url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
        metadata = session.get(
            metadata_url,
            params={"fields": "name,mimeType"},
            headers=headers,
            timeout=DOWNLOAD_TIMEOUT_SECONDS,
            verify=VERIFY_SSL_CERTIFICATES,
        )
        metadata.raise_for_status()
        details = metadata.json()
        filename = safe_filename(details.get("name", fallback), fallback)
        mime_type = details.get("mimeType", "")
        if mime_type == "application/vnd.google-apps.spreadsheet":
            url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export"
            params = {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
            if not Path(filename).suffix:
                filename += ".xlsx"
        else:
            url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
            params = {"alt": "media"}
        response = session.get(
            url,
            params=params,
            headers=headers,
            stream=True,
            timeout=DOWNLOAD_TIMEOUT_SECONDS,
            verify=VERIFY_SSL_CERTIFICATES,
        )
        return stream_response_to_file(response, destination_dir / filename)

    parsed = urlparse(link)
    if "spreadsheets" in parsed.path:
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/export"
        response = session.get(
            url,
            params={"format": "xlsx"},
            stream=True,
            timeout=DOWNLOAD_TIMEOUT_SECONDS,
            verify=VERIFY_SSL_CERTIFICATES,
        )
        return stream_response_to_file(response, destination_dir / f"{fallback}.xlsx")

    response = session.get(
        "https://drive.google.com/uc",
        params={"export": "download", "id": file_id},
        stream=True,
        timeout=DOWNLOAD_TIMEOUT_SECONDS,
        verify=VERIFY_SSL_CERTIFICATES,
    )
    content_type = response.headers.get("Content-Type", "").casefold()
    if "text/html" in content_type:
        page = response.text
        parser = DownloadConfirmationParser()
        parser.feed(page)
        if not parser.action:
            raise RuntimeError(
                "Google Drive returned a sign-in/permission page. Make the file link-accessible "
                "or configure a Drive access token/service account."
            )
        response = session.get(
            urljoin(response.url, parser.action),
            params=parser.inputs,
            stream=True,
            timeout=DOWNLOAD_TIMEOUT_SECONDS,
            verify=VERIFY_SSL_CERTIFICATES,
        )
    filename = filename_from_response(response, fallback)
    return stream_response_to_file(response, destination_dir / filename)


def download_file(
    session: requests.Session, link: str, destination_dir: Path, index: int
) -> Path:
    if os.environ.get("NCVET_WEB_MODE") == "1":
        from web_downloads import validate_url
        validate_url(link)
    destination_dir.mkdir(parents=True, exist_ok=True)
    # urlparse treats a Windows drive letter as a URL scheme ("C:"). Check
    # Windows/UNC paths before parsing web URLs so local test or network-share
    # inputs work correctly.
    if re.match(r"^[A-Za-z]:[\\/]", link) or link.startswith("\\\\"):
        source = Path(link)
        if not source.is_file():
            raise RuntimeError(f"Local file does not exist: {source}")
        target = unique_path(destination_dir / safe_filename(source.name, f"file_{index}"))
        shutil.copy2(source, target)
        return target
    parsed = urlparse(link)
    if parsed.scheme.casefold() == "file":
        source = Path(unquote(parsed.path.lstrip("/") if os.name == "nt" else parsed.path))
        if os.name == "nt" and re.match(r"^[A-Za-z]:", unquote(parsed.netloc + parsed.path)):
            source = Path(unquote(parsed.netloc + parsed.path))
        if not source.is_file():
            raise RuntimeError(f"Local file does not exist: {source}")
        target = unique_path(destination_dir / safe_filename(source.name, f"file_{index}"))
        shutil.copy2(source, target)
        return target
    if not parsed.scheme:
        source = Path(link)
        if source.is_file():
            target = unique_path(destination_dir / safe_filename(source.name, f"file_{index}"))
            shutil.copy2(source, target)
            return target
        raise RuntimeError(f"Local file does not exist: {source}")
    if parsed.scheme.casefold() not in {"http", "https"}:
        raise RuntimeError(f"Unsupported link type: {parsed.scheme}")

    hostname = parsed.hostname.casefold() if parsed.hostname else ""
    if hostname.endswith("drive.google.com") or hostname.endswith("docs.google.com"):
        return download_google_drive(session, link, destination_dir, index)

    response = session.get(
        link,
        stream=True,
        timeout=DOWNLOAD_TIMEOUT_SECONDS,
        verify=VERIFY_SSL_CERTIFICATES,
    )
    filename = filename_from_response(response, f"download_{index}")
    return stream_response_to_file(response, destination_dir / filename)


def iter_xlsx_rows(path: Path) -> Iterator[tuple[str, list[list[Any]]]]:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        for sheet in workbook.worksheets:
            rows: list[list[Any]] = []
            # Some producer applications serialize an entirely empty worksheet
            # without a dimension, which openpyxl exposes as None.
            max_col = min(
                sheet.max_column or MAX_COLUMNS_TO_SCAN, MAX_COLUMNS_TO_SCAN
            )
            max_row = min(sheet.max_row or HEADER_SCAN_ROWS, HEADER_SCAN_ROWS)
            for row in sheet.iter_rows(max_row=max_row, max_col=max_col, values_only=True):
                rows.append(list(row))
            yield sheet.title, rows
    finally:
        workbook.close()


def iter_xlsb_rows(path: Path) -> Iterator[tuple[str, list[list[Any]]]]:
    try:
        from pyxlsb import open_workbook
    except ImportError as exc:
        raise RuntimeError(
            "Reading .xlsb files requires pyxlsb. Run: pip install -r requirements.txt"
        ) from exc
    with open_workbook(path) as workbook:
        for sheet_name in workbook.sheets:
            rows: list[list[Any]] = []
            with workbook.get_sheet(sheet_name) as sheet:
                for row_index, row in enumerate(sheet.rows()):
                    if row_index >= HEADER_SCAN_ROWS:
                        break
                    rows.append([cell.v for cell in row[:MAX_COLUMNS_TO_SCAN]])
            yield sheet_name, rows


def iter_xls_rows(path: Path) -> Iterator[tuple[str, list[list[Any]]]]:
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError(
            "Reading legacy .xls files requires pandas and xlrd. Run: pip install -r requirements.txt"
        ) from exc
    try:
        workbook = pd.ExcelFile(path, engine="xlrd")
    except ImportError as exc:
        raise RuntimeError(
            "Reading legacy .xls files requires xlrd. Run: pip install -r requirements.txt"
        ) from exc
    for sheet_name in workbook.sheet_names:
        frame = workbook.parse(sheet_name, header=None, nrows=HEADER_SCAN_ROWS)
        frame = frame.iloc[:, :MAX_COLUMNS_TO_SCAN]
        yield sheet_name, frame.where(frame.notna(), None).values.tolist()


def validate_workbook(path: Path, required_columns: list[str]) -> FileValidation:
    result = FileValidation(file_path=path)
    required_normalized = [normalize_header(value) for value in required_columns]
    required_set = set(required_normalized)
    try:
        extension = sniff_excel_extension(path)
        if extension in {".xlsx", ".xlsm"}:
            sheets = iter_xlsx_rows(path)
        elif extension == ".xlsb":
            sheets = iter_xlsb_rows(path)
        elif extension == ".xls":
            sheets = iter_xls_rows(path)
        else:
            raise RuntimeError(f"Unsupported file type '{path.suffix or extension}'. Expected .xlsx or .xlsb.")

        best_score = -1
        best_headers: set[str] = set()
        for sheet_name, rows in sheets:
            for row_number, row in enumerate(rows, start=1):
                actual = {normalize_header(value) for value in row if normalize_header(value)}
                score = len(required_set.intersection(actual))
                if score > best_score:
                    best_score = score
                    best_headers = actual
                    result.sheet_name = sheet_name
                    result.header_row = row_number
                    result.matched_column_count = score
        result.missing_columns = [
            original
            for original, normalized in zip(required_columns, required_normalized)
            if normalized not in best_headers
        ]
        return result
    except Exception as exc:
        result.error = f"{type(exc).__name__}: {exc}"
        return result


def aligned_values(
    row: list[Any], header_indexes: list[int | None]
) -> list[Any]:
    """Return one source row reordered to the DCF column sequence."""
    return [
        row[index] if index is not None and index < len(row) else None
        for index in header_indexes
    ]


def required_header_indexes(
    header_row: list[Any], required_columns: list[str]
) -> list[int | None]:
    actual = header_map(header_row)
    return [actual.get(normalize_header(column)) for column in required_columns]


def row_has_data(row: Iterable[Any]) -> bool:
    return any(normalize_text(value) for value in row)


def iter_aligned_data_rows(
    path: Path,
    validation: FileValidation,
    required_columns: list[str],
) -> Iterator[tuple[int, list[Any]]]:
    """Yield (source row number, DCF-aligned values) from a validated file."""
    if not validation.sheet_name or not validation.header_row:
        raise RuntimeError("The workbook's data header row could not be identified.")

    extension = sniff_excel_extension(path)
    if extension in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if validation.sheet_name not in workbook.sheetnames:
                raise RuntimeError(
                    f"Validated sheet no longer exists: {validation.sheet_name}"
                )
            sheet = workbook[validation.sheet_name]
            max_col = min(
                sheet.max_column or MAX_COLUMNS_TO_SCAN, MAX_COLUMNS_TO_SCAN
            )
            header = list(
                next(
                    sheet.iter_rows(
                        min_row=validation.header_row,
                        max_row=validation.header_row,
                        max_col=max_col,
                        values_only=True,
                    )
                )
            )
            indexes = required_header_indexes(header, required_columns)
            for row_number, row in enumerate(
                sheet.iter_rows(
                    min_row=validation.header_row + 1,
                    max_col=max_col,
                    values_only=True,
                ),
                start=validation.header_row + 1,
            ):
                values = list(row)
                if row_has_data(values):
                    yield row_number, aligned_values(values, indexes)
        finally:
            workbook.close()
        return

    if extension == ".xlsb":
        try:
            from pyxlsb import open_workbook
        except ImportError as exc:
            raise RuntimeError(
                "Reading .xlsb files requires pyxlsb. Run: pip install -r requirements.txt"
            ) from exc
        with open_workbook(path) as workbook:
            if validation.sheet_name not in workbook.sheets:
                raise RuntimeError(
                    f"Validated sheet no longer exists: {validation.sheet_name}"
                )
            with workbook.get_sheet(validation.sheet_name) as sheet:
                indexes: list[int | None] | None = None
                for row_number, row in enumerate(sheet.rows(), start=1):
                    values = [cell.v for cell in row[:MAX_COLUMNS_TO_SCAN]]
                    if row_number == validation.header_row:
                        indexes = required_header_indexes(values, required_columns)
                        continue
                    if row_number <= validation.header_row:
                        continue
                    if indexes is None:
                        raise RuntimeError("Could not read the validated .xlsb header row.")
                    if row_has_data(values):
                        yield row_number, aligned_values(values, indexes)
        return

    if extension == ".xls":
        try:
            import pandas as pd
        except ImportError as exc:
            raise RuntimeError(
                "Reading legacy .xls files requires pandas and xlrd. Run: pip install -r requirements.txt"
            ) from exc
        try:
            frame = pd.read_excel(
                path,
                sheet_name=validation.sheet_name,
                header=None,
                engine="xlrd",
            )
        except ImportError as exc:
            raise RuntimeError(
                "Reading legacy .xls files requires xlrd. Run: pip install -r requirements.txt"
            ) from exc
        frame = frame.iloc[:, :MAX_COLUMNS_TO_SCAN].where(frame.notna(), None)
        rows = frame.values.tolist()
        if validation.header_row > len(rows):
            raise RuntimeError("The validated .xls header row is outside the worksheet data.")
        indexes = required_header_indexes(
            rows[validation.header_row - 1], required_columns
        )
        for row_number, values in enumerate(
            rows[validation.header_row :], start=validation.header_row + 1
        ):
            if row_has_data(values):
                yield row_number, aligned_values(values, indexes)
        return

    raise RuntimeError(
        f"Unsupported file type '{path.suffix or extension}'. Expected .xlsx or .xlsb."
    )


def evaluate_response(
    response: ResponseRecord,
    month: ReportingMonth,
    agency: Agency,
    downloads_root: Path,
    session: requests.Session,
    required_columns: list[str],
) -> StatusEvidence:
    if is_no_assessment(response.conducted_raw):
        return StatusEvidence(status=STATUS_NO_ASSESSMENTS, response=response)
    if not response.links:
        return StatusEvidence(
            status=STATUS_FILE_ERROR,
            response=response,
            errors=["The response indicates assessments/data submission but contains no downloadable link."],
        )

    agency_folder = downloads_root / month.folder_name / safe_filename(agency.name, "Agency")
    evidence = StatusEvidence(status=STATUS_SUBMITTED, response=response)
    for index, link in enumerate(response.links, start=1):
        try:
            downloaded = download_file(session, link, agency_folder, index)
            evidence.downloaded_files.append(downloaded)
            validation = validate_workbook(downloaded, required_columns)
            evidence.validations.append(validation)
            if validation.error:
                evidence.errors.append(f"{downloaded.name}: {validation.error}")
            evidence.missing_columns.extend(validation.missing_columns)
        except Exception as exc:
            evidence.errors.append(f"Link {index}: {type(exc).__name__}: {exc}")

    if evidence.errors:
        evidence.status = STATUS_FILE_ERROR
    elif evidence.missing_columns:
        evidence.status = STATUS_NOT_IN_FORMAT
    else:
        evidence.status = STATUS_SUBMITTED
    evidence.missing_columns = list(dict.fromkeys(evidence.missing_columns))
    return evidence


def write_only_cell(sheet, value: Any, header: bool = False) -> WriteOnlyCell:
    cell = WriteOnlyCell(sheet, value=text_for_excel(value))
    if header:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    else:
        cell.font = Font(name="Aptos", size=10, color="1F1F1F")
        cell.alignment = Alignment(vertical="top", wrap_text=False)
        if isinstance(value, datetime):
            cell.number_format = "yyyy-mm-dd hh:mm:ss"
        elif isinstance(value, date):
            cell.number_format = "yyyy-mm-dd"
    return cell


def build_consolidated_master(
    master_path: Path,
    agencies: list[Agency],
    months: list[ReportingMonth],
    evidence_by_key: dict[tuple[str, ReportingMonth], StatusEvidence],
    required_columns: list[str],
    run_folder: Path,
) -> int:
    """Combine every parseable downloaded workbook into one DCF-aligned sheet."""
    workbook = Workbook(write_only=True)
    workbook.creator = "NCVET AA Monitoring Tool"
    workbook.title = "Consolidated Assessment Agency Data"

    metadata_headers = [
        "Reporting Month",
        "Assessment Agency (Master)",
        "Source File Format Status",
        "Source File",
        "Source Sheet",
        "Source Row",
    ]
    master_sheet = workbook.create_sheet("Master Data")
    master_sheet.sheet_view.showGridLines = False
    master_sheet.freeze_panes = "A2"
    master_sheet.row_dimensions[1].height = 54
    master_sheet.column_dimensions["A"].width = 16
    master_sheet.column_dimensions["B"].width = 42
    master_sheet.column_dimensions["C"].width = 24
    master_sheet.column_dimensions["D"].width = 48
    master_sheet.column_dimensions["E"].width = 24
    master_sheet.column_dimensions["F"].width = 12
    for column in range(7, 7 + len(required_columns)):
        master_sheet.column_dimensions[get_column_letter(column)].width = 22
    master_sheet.append(
        [
            write_only_cell(master_sheet, value, header=True)
            for value in metadata_headers + required_columns
        ]
    )

    log_headers = [
        "Reporting Month",
        "Assessment Agency (Master)",
        "Source File",
        "Source Sheet",
        "Header Row",
        "Matched DCF Columns",
        "Missing DCF Columns",
        "Rows Consolidated",
        "Result / Error",
    ]
    log_rows: list[list[Any]] = []
    total_rows = 0

    for month in months:
        for agency in agencies:
            evidence = evidence_by_key[(agency.normalized_name, month)]
            for validation in evidence.validations:
                relative_file = str(validation.file_path.relative_to(run_folder))
                if validation.error:
                    log_rows.append(
                        [
                            month.label,
                            agency.name,
                            relative_file,
                            validation.sheet_name,
                            validation.header_row,
                            validation.matched_column_count,
                            "\n".join(validation.missing_columns),
                            0,
                            validation.error,
                        ]
                    )
                    continue

                file_status = (
                    STATUS_NOT_IN_FORMAT
                    if validation.missing_columns
                    else STATUS_SUBMITTED
                )
                rows_added = 0
                try:
                    for source_row, aligned_row in iter_aligned_data_rows(
                        validation.file_path, validation, required_columns
                    ):
                        values = [
                            month.label,
                            agency.name,
                            file_status,
                            relative_file,
                            validation.sheet_name,
                            source_row,
                        ] + aligned_row
                        master_sheet.append(
                            [write_only_cell(master_sheet, value) for value in values]
                        )
                        rows_added += 1
                        total_rows += 1
                    validation.rows_consolidated = rows_added
                    result_note = (
                        "Included; missing DCF fields left blank"
                        if validation.missing_columns
                        else "Included"
                    )
                except Exception as exc:
                    validation.consolidation_error = f"{type(exc).__name__}: {exc}"
                    evidence.errors.append(
                        f"{validation.file_path.name}: consolidation failed: "
                        f"{validation.consolidation_error}"
                    )
                    evidence.status = STATUS_FILE_ERROR
                    result_note = validation.consolidation_error
                log_rows.append(
                    [
                        month.label,
                        agency.name,
                        relative_file,
                        validation.sheet_name,
                        validation.header_row,
                        validation.matched_column_count,
                        "\n".join(validation.missing_columns),
                        rows_added,
                        result_note,
                    ]
                )

    last_master_column = get_column_letter(len(metadata_headers) + len(required_columns))
    master_sheet.auto_filter.ref = f"A1:{last_master_column}{total_rows + 1}"

    log_sheet = workbook.create_sheet("Consolidation Log")
    log_sheet.sheet_view.showGridLines = False
    log_sheet.freeze_panes = "A2"
    log_sheet.row_dimensions[1].height = 38
    log_widths = [16, 42, 48, 24, 12, 22, 70, 20, 70]
    for column, width in enumerate(log_widths, start=1):
        log_sheet.column_dimensions[get_column_letter(column)].width = width
    log_sheet.append(
        [write_only_cell(log_sheet, value, header=True) for value in log_headers]
    )
    for row in log_rows:
        log_sheet.append([write_only_cell(log_sheet, value) for value in row])
    log_sheet.auto_filter.ref = f"A1:I{len(log_rows) + 1}"

    master_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(master_path)
    return total_rows


def add_table(sheet, reference: str, name: str) -> None:
    table = Table(displayName=name, ref=reference)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def text_for_excel(value: Any, limit: int = 32000) -> Any:
    if isinstance(value, str):
        value = EXCEL_ILLEGAL_CHARS.sub("", value)
    if isinstance(value, str) and len(value) > limit:
        return value[: limit - 16] + " … [truncated]"
    return value


def build_status_report(
    report_path: Path,
    consolidated_master_path: Path,
    agencies: list[Agency],
    months: list[ReportingMonth],
    evidence_by_key: dict[tuple[str, ReportingMonth], StatusEvidence],
    records: list[ResponseRecord],
    required_columns: list[str],
    run_folder: Path,
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    workbook.creator = "NCVET AA Monitoring Tool"
    workbook.title = "Assessment Agency Submission Status"

    navy = "17365D"
    white = "FFFFFF"
    light_border = Side(style="thin", color="D9E2F3")
    medium_border = Side(style="medium", color=navy)
    header_fill = PatternFill("solid", fgColor=navy)
    header_font = Font(name="Aptos", size=10, bold=True, color=white)
    body_font = Font(name="Aptos", size=10, color="1F1F1F")

    status_sheet = workbook.create_sheet("Status")
    status_sheet.sheet_view.showGridLines = False
    last_column = 3 + len(months)
    last_column_letter = get_column_letter(last_column)
    status_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    title_cell = status_sheet.cell(1, 1, "Assessment Agency Submission Status")
    title_cell.fill = header_fill
    title_cell.font = Font(name="Aptos Display", size=16, bold=True, color=white)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    status_sheet.row_dimensions[1].height = 30

    summary_header_row = 3
    summary_headers = ["Status", "All periods"] + [month.label for month in months]
    for column, value in enumerate(summary_headers, start=1):
        cell = status_sheet.cell(summary_header_row, column, value)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=medium_border)

    status_start_row = 12
    status_end_row = status_start_row + len(agencies) - 1
    first_month_col = 4
    last_month_col = first_month_col + len(months) - 1

    summary_labels = ["Total"] + STATUS_ORDER
    for offset, label in enumerate(summary_labels, start=1):
        row = summary_header_row + offset
        status_sheet.cell(row, 1, label)
        status_sheet.cell(row, 1).font = Font(name="Aptos", size=10, bold=True)
        fill_color = "E2F0D9" if label == "Total" else STATUS_COLORS[label]
        for column in range(1, 3 + len(months)):
            status_sheet.cell(row, column).fill = PatternFill("solid", fgColor=fill_color)
            status_sheet.cell(row, column).border = Border(bottom=light_border)
        if label == "Total":
            if months:
                status_sheet.cell(
                    row,
                    2,
                    f"=COUNTA({get_column_letter(first_month_col)}{status_start_row}:{get_column_letter(last_month_col)}{status_end_row})",
                )
            else:
                status_sheet.cell(row, 2, 0)
        else:
            status_sheet.cell(
                row,
                2,
                f'=COUNTIF({get_column_letter(first_month_col)}{status_start_row}:{get_column_letter(last_month_col)}{status_end_row},A{row})',
            )
        for month_index, _month in enumerate(months):
            column = first_month_col + month_index
            column_letter = get_column_letter(column)
            if label == "Total":
                formula = f"=COUNTA({column_letter}{status_start_row}:{column_letter}{status_end_row})"
            else:
                formula = f'=COUNTIF({column_letter}{status_start_row}:{column_letter}{status_end_row},$A{row})'
            status_sheet.cell(row, 3 + month_index, formula)

    table_headers = ["S.No.", f"Assessment Agency Name ({len(agencies)})", "Jurisdiction"] + [
        month.label for month in months
    ]
    for column, value in enumerate(table_headers, start=1):
        cell = status_sheet.cell(11, column, value)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=medium_border)
    status_sheet.row_dimensions[11].height = 28

    for row_offset, agency in enumerate(agencies):
        row = status_start_row + row_offset
        status_sheet.cell(row, 1, agency.serial_no)
        status_sheet.cell(row, 2, agency.name)
        status_sheet.cell(row, 3, agency.jurisdiction)
        for month_index, month in enumerate(months):
            status = evidence_by_key[(agency.normalized_name, month)].status
            cell = status_sheet.cell(row, first_month_col + month_index, status)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for column in range(1, last_column + 1):
            cell = status_sheet.cell(row, column)
            cell.font = body_font
            cell.border = Border(bottom=light_border)
            if column in (1, 3):
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif column == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    if months:
        status_range = f"D{status_start_row}:{get_column_letter(last_month_col)}{status_end_row}"
        anchor = f"D{status_start_row}"
        for status, color in STATUS_COLORS.items():
            status_sheet.conditional_formatting.add(
                status_range,
                FormulaRule(
                    formula=[f'{anchor}="{status}"'],
                    fill=PatternFill("solid", fgColor=color),
                ),
            )

    status_sheet.freeze_panes = f"D{status_start_row}"
    status_sheet.auto_filter.ref = f"A11:{last_column_letter}{status_end_row}"
    # Wide enough for the summary labels (for example "Data Not Submitted").
    status_sheet.column_dimensions["A"].width = 22
    status_sheet.column_dimensions["B"].width = 52
    status_sheet.column_dimensions["C"].width = 30
    for column in range(4, last_column + 1):
        status_sheet.column_dimensions[get_column_letter(column)].width = 17
    status_sheet.sheet_properties.pageSetUpPr.fitToPage = True
    status_sheet.page_setup.fitToWidth = 1
    status_sheet.page_setup.fitToHeight = 0
    status_sheet.print_title_rows = "1:11"

    detail_sheet = workbook.create_sheet("Validation Details")
    detail_sheet.sheet_view.showGridLines = False
    detail_headers = [
        "S.No.",
        "Assessment Agency Name",
        "Jurisdiction",
        "Month",
        "Status",
        "Response Timestamp",
        "Conducted Answer",
        "Source Link(s)",
        "Downloaded File(s)",
        "Matched Header Sheet/Row",
        "Rows Consolidated",
        "Missing DCF Column(s)",
        "Error / Note",
    ]
    detail_sheet.append(detail_headers)
    for agency in agencies:
        for month in months:
            evidence = evidence_by_key[(agency.normalized_name, month)]
            response = evidence.response
            matches = []
            for validation in evidence.validations:
                if validation.sheet_name:
                    matches.append(
                        f"{validation.file_path.name}: {validation.sheet_name}!row {validation.header_row} "
                        f"({validation.matched_column_count}/{len(required_columns)} headers)"
                    )
            detail_sheet.append(
                [
                    agency.serial_no,
                    agency.name,
                    agency.jurisdiction,
                    month.label,
                    evidence.status,
                    response.timestamp if response and response.timestamp else normalize_text(response.timestamp_raw) if response else "",
                    response.conducted_raw if response else "",
                    "\n".join(response.links) if response else "",
                    "\n".join(str(path.relative_to(run_folder)) for path in evidence.downloaded_files),
                    "\n".join(matches),
                    sum(validation.rows_consolidated for validation in evidence.validations),
                    "\n".join(evidence.missing_columns),
                    "\n".join(evidence.errors) or (response.response_note if response else "No response found for this agency/month"),
                ]
            )
    for cell in detail_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    detail_sheet.freeze_panes = "A2"
    detail_sheet.auto_filter.ref = f"A1:M{detail_sheet.max_row}"
    widths = [10, 42, 28, 12, 20, 21, 20, 46, 46, 42, 18, 70, 70]
    for column, width in enumerate(widths, start=1):
        detail_sheet.column_dimensions[get_column_letter(column)].width = width
    for row in detail_sheet.iter_rows(min_row=2):
        for cell in row:
            cell.value = text_for_excel(cell.value)
            cell.font = body_font
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    if detail_sheet.max_row >= 2:
        add_table(detail_sheet, f"A1:M{detail_sheet.max_row}", "ValidationDetailsTable")

    response_sheet = workbook.create_sheet("Response Log")
    response_sheet.sheet_view.showGridLines = False
    response_headers = [
        "Response Row",
        "Timestamp",
        "Submitted Agency Name",
        "Matched Master Agency",
        "Match Method",
        "Match Score",
        "Submitted Month",
        "Parsed Month",
        "Conducted Answer",
        "Link Count",
        "Selected Latest?",
        "Note",
    ]
    response_sheet.append(response_headers)
    for record in records:
        response_sheet.append(
            [
                record.row_number,
                record.timestamp or normalize_text(record.timestamp_raw),
                record.agency_raw,
                record.matched_agency.name if record.matched_agency else "",
                record.match_method,
                record.match_score,
                normalize_text(record.month_raw),
                record.month.label if record.month else "",
                record.conducted_raw,
                len(record.links),
                "Yes" if record.selected else "No",
                record.response_note,
            ]
        )
    for cell in response_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    response_sheet.freeze_panes = "A2"
    response_sheet.auto_filter.ref = f"A1:L{max(response_sheet.max_row, 1)}"
    response_widths = [14, 21, 42, 42, 16, 14, 18, 14, 22, 12, 16, 58]
    for column, width in enumerate(response_widths, start=1):
        response_sheet.column_dimensions[get_column_letter(column)].width = width
    if response_sheet.max_row >= 2:
        for cell in response_sheet["F"][1:]:
            cell.number_format = "0.0%"
        add_table(response_sheet, f"A1:L{response_sheet.max_row}", "ResponseLogTable")

    required_sheet = workbook.create_sheet("Required Columns")
    required_sheet.sheet_view.showGridLines = False
    required_sheet.append(["S.No.", "Required DCF Column"])
    for index, column_name in enumerate(required_columns, start=1):
        required_sheet.append([index, column_name])
    for cell in required_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    required_sheet.column_dimensions["A"].width = 10
    required_sheet.column_dimensions["B"].width = 78
    required_sheet.freeze_panes = "A2"
    add_table(required_sheet, f"A1:B{required_sheet.max_row}", "RequiredColumnsTable")

    about_sheet = workbook.create_sheet("About")
    about_sheet.sheet_view.showGridLines = False
    about_rows = [
        ["NCVET Assessment Agency Monitoring Report"],
        ["Run folder", str(run_folder)],
        ["Consolidated master", str(consolidated_master_path)],
        ["Rule", "The latest response by Timestamp is used for each matched agency and month."],
        ["Master data rule", "Every parseable data row is aligned to the DCF columns. Missing fields in Not in Format files are left blank."],
        [STATUS_SUBMITTED, "All linked files downloaded, opened, and contained every required DCF column."],
        [STATUS_NOT_IN_FORMAT, "At least one linked file opened but one or more required DCF columns were missing."],
        [STATUS_NO_ASSESSMENTS, "The latest form response said assessments were not conducted."],
        [STATUS_DATA_NOT_SUBMITTED, "No matched form response was found for that agency/month."],
        [STATUS_FILE_ERROR, "A link was absent, download failed, file was inaccessible, or workbook could not be parsed."],
        ["Google Drive", "Private Drive uploads require a configured OAuth token or a service account with access."],
    ]
    for row in about_rows:
        about_sheet.append(row)
    about_sheet.merge_cells("A1:B1")
    about_sheet["A1"].fill = header_fill
    about_sheet["A1"].font = Font(name="Aptos Display", size=15, bold=True, color=white)
    about_sheet["A1"].alignment = Alignment(vertical="center")
    for row in range(2, about_sheet.max_row + 1):
        about_sheet.cell(row, 1).font = Font(name="Aptos", size=10, bold=True)
        about_sheet.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="top")
    about_sheet.column_dimensions["A"].width = 25
    about_sheet.column_dimensions["B"].width = 100
    about_sheet.row_dimensions[1].height = 28

    report_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(report_path)


def resolve_reporting_months(records: list[ResponseRecord]) -> list[ReportingMonth]:
    if REPORTING_MONTHS:
        months: list[ReportingMonth] = []
        for value in REPORTING_MONTHS:
            parsed = parse_reporting_month(value, default_year=DEFAULT_REPORTING_YEAR)
            if not parsed:
                raise ValueError(
                    f"Could not parse REPORTING_MONTHS value {value!r}. Use a value such as Apr-26."
                )
            if parsed not in months:
                months.append(parsed)
        return months
    inferred = sorted({record.month for record in records if record.month})
    if not inferred:
        raise ValueError(
            "No reporting month could be inferred. Add years to the Month values, ensure Timestamp is valid, "
            "or set REPORTING_MONTHS at the top of the script."
        )
    return inferred


def validate_input_paths(responses: Path, dcf: Path, master: Path) -> None:
    for label, path in (
        ("FORM_RESPONSES_PATH", responses),
        ("DCF_TEMPLATE_PATH", dcf),
        ("AA_CATEGORIES_PATH", master),
    ):
        if not path.is_file():
            raise FileNotFoundError(f"{label} does not point to a file: {path}")


def run(
    responses_path: Path,
    dcf_path: Path,
    master_path: Path,
    output_root: Path,
) -> tuple[Path, Path]:
    validate_input_paths(responses_path, dcf_path, master_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_folder = output_root / f"run_{timestamp}" if CREATE_TIMESTAMPED_RUN_FOLDER else output_root
    run_folder.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(run_folder / "processing.log", encoding="utf-8"),
        ],
        force=True,
    )
    logging.info("Loading agency master: %s", master_path)
    agencies = load_agency_master(master_path)
    logging.info("Loaded %d agencies", len(agencies))
    required_columns = load_required_columns(dcf_path)
    logging.info("Loaded %d required DCF columns", len(required_columns))
    records = read_form_responses(responses_path)
    logging.info("Loaded %d form responses", len(records))
    match_agencies(records, agencies)
    months = resolve_reporting_months(records)
    logging.info("Reporting months: %s", ", ".join(month.label for month in months))
    selected = select_latest_responses(records, months)

    session = requests.Session()
    if os.environ.get("NCVET_WEB_MODE") == "1":
        from web_downloads import GoogleDownloadAdapter
        session.mount("https://", GoogleDownloadAdapter())
        session.mount("http://", GoogleDownloadAdapter())
    session.verify = VERIFY_SSL_CERTIFICATES
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 NCVET-AA-Monitor/1.0",
            "Accept": "*/*",
        }
    )
    evidence_by_key: dict[tuple[str, ReportingMonth], StatusEvidence] = {}
    downloads_root = run_folder / "downloaded_files"
    for month in months:
        for agency in agencies:
            key = (agency.normalized_name, month)
            response = selected.get(key)
            if response is None:
                evidence = StatusEvidence(status=STATUS_DATA_NOT_SUBMITTED)
            else:
                logging.info("Processing %s | %s", agency.name, month.label)
                evidence = evaluate_response(
                    response,
                    month,
                    agency,
                    downloads_root,
                    session,
                    required_columns,
                )
            evidence_by_key[key] = evidence

    consolidated_master_path = run_folder / CONSOLIDATED_MASTER_FILENAME
    logging.info("Combining AA rows into: %s", consolidated_master_path)
    consolidated_row_count = build_consolidated_master(
        consolidated_master_path,
        agencies,
        months,
        evidence_by_key,
        required_columns,
        run_folder,
    )
    logging.info("Consolidated %d data rows", consolidated_row_count)

    report_path = run_folder / "AA_Status_Report.xlsx"
    build_status_report(
        report_path,
        consolidated_master_path,
        agencies,
        months,
        evidence_by_key,
        records,
        required_columns,
        run_folder,
    )

    counts = {status: 0 for status in STATUS_ORDER}
    for evidence in evidence_by_key.values():
        counts[evidence.status] += 1
    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "responses_file": str(responses_path),
        "dcf_template": str(dcf_path),
        "agency_master": str(master_path),
        "run_folder": str(run_folder),
        "status_report": str(report_path),
        "consolidated_master": str(consolidated_master_path),
        "consolidated_data_row_count": consolidated_row_count,
        "agency_count": len(agencies),
        "required_column_count": len(required_columns),
        "response_count": len(records),
        "months": [month.label for month in months],
        "status_counts_across_agency_months": counts,
    }
    (run_folder / "run_manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    logging.info("Complete. Status report: %s", report_path)
    logging.info("Complete. Consolidated master: %s", consolidated_master_path)
    return report_path, consolidated_master_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Download AA form submissions, validate DCF columns, combine AA rows, "
            "and create a status report."
        )
    )
    parser.add_argument("--responses", type=Path, default=FORM_RESPONSES_PATH)
    parser.add_argument("--dcf", type=Path, default=DCF_TEMPLATE_PATH)
    parser.add_argument("--master", type=Path, default=AA_CATEGORIES_PATH)
    parser.add_argument("--output", type=Path, default=OUTPUT_ROOT)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        report, consolidated_master = run(
            args.responses, args.dcf, args.master, args.output
        )
    except Exception as exc:
        logging.exception("Processing failed")
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    print(f"\nDone. Status report: {report}")
    print(f"Done. Consolidated master: {consolidated_master}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
