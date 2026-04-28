"""
Phase 1 - ComplySci Marketing Material Request matcher

What this script does:
1. Opens ComplySci and logs in.
2. Navigates to Marketing Material Requests.
3. Switches the grid to Processed.
4. Applies the Username filter using a firm keyword.
5. Reads approved rows from the Excel sheet.
6. Scrapes approved rows from the website.
7. Matches Excel rows to website rows by Employee + Processed Date + Approved status.
8. Saves the phase-1 output to JSON so we can build phase 2 on top of it.

What this script does NOT do yet:
- download attached documents
- click Export to PDF
- rename files
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import time
import traceback
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable

import openpyxl
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver import ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================================
# CONFIG
# ============================================================================

ENV_FILE_NAME = ".env"


def load_local_env(env_path: str | Path) -> dict[str, str]:
    env_file = Path(env_path)
    loaded: dict[str, str] = {}

    if not env_file.exists():
        return loaded

    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()

        if value and len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]

        loaded[key] = value
        os.environ[key] = value

    return loaded


load_local_env(Path(__file__).with_name(ENV_FILE_NAME))


EMAIL = os.getenv("EMAIL", "").strip()
PASSWORD = os.getenv("PASSWORD", "").strip()

FILTER_KEYWORD = os.getenv("FILTER_KEYWORD", "").strip()

EXCEL_PATH = os.getenv("EXCEL_PATH", "").strip()
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "").strip()


# ============================================================================
# CONSTANTS
# ============================================================================

BASE_URL = "https://truindependence.complysci.com"
REQUESTS_URL = f"{BASE_URL}/PreClear/MarketingApproverPreclearanceRequests"
DASHBOARD_URL = f"{BASE_URL}/Home/ComplianceOfficerDashboardNew"

WAIT_TIMEOUT = 30
ARTIFACTS_DIRNAME = "_phase1_artifacts"
JSON_OUTPUT_NAME = "phase1_matches.json"
DOWNLOAD_SUMMARY_NAME = "phase2_download_summary.json"
DOWNLOAD_ISSUES_NAME = "phase2_download_issues.txt"
DOWNLOAD_TIMEOUT = 45


# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("complysci_downloader.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ============================================================================
# DATA MODELS
# ============================================================================

@dataclass
class ExcelRecord:
    firm: str
    employee: str
    processed_date: date
    title: str
    status: str


@dataclass
class WebRecord:
    employee: str
    username: str
    processed_date: date
    processed_date_text: str
    status: str
    title: str
    row_page: int
    row_index: int
    view_href: str | None
    indirect_request_id: str | None


@dataclass
class MatchRecord:
    firm: str
    employee: str
    processed_date_excel: str
    processed_date_web: str
    username: str
    title: str
    status: str
    row_page: int
    row_index: int
    view_href: str | None
    indirect_request_id: str | None


@dataclass
class DownloadOutcome:
    firm: str
    employee: str
    processed_date: str
    indirect_request_id: str | None
    approval_file: str | None
    doc_files: list[str]


@dataclass
class Phase2Issue:
    recorded_at: str
    issue_type: str
    firm: str
    employee: str
    processed_date: str
    indirect_request_id: str | None
    details: str


# ============================================================================
# GENERIC HELPERS
# ============================================================================

def normalise_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def canonical_person_name(value: str) -> str:
    return normalise_space(value).casefold()


def parse_date_value(raw) -> date | None:
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw

    raw_text = normalise_space(str(raw))
    if not raw_text:
        return None

    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw_text, fmt).date()
        except ValueError:
            continue

    return None


def parse_date_from_web_text(raw_text: str) -> date | None:
    match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", raw_text or "")
    if not match:
        return None
    return datetime.strptime(match.group(1), "%m/%d/%Y").date()


def format_web_datetime(raw) -> str:
    if raw is None:
        return ""
    if isinstance(raw, str):
        return normalise_space(raw)
    if isinstance(raw, (int, float)):
        try:
            raw = datetime.fromtimestamp(raw / 1000)
        except Exception:
            return str(raw)
    if isinstance(raw, datetime):
        return raw.strftime("%-m/%-d/%Y %-I:%M:%S %p").replace("AM", "AM (PDT)").replace("PM", "PM (PDT)")
    return normalise_space(str(raw))


def ensure_dir(path: str | Path) -> Path:
    target = Path(path)
    target.mkdir(parents=True, exist_ok=True)
    return target


def append_phase2_issue(
    issue_log_path: Path,
    issues: list[Phase2Issue],
    match: MatchRecord,
    issue_type: str,
    details: str,
) -> None:
    issue = Phase2Issue(
        recorded_at=datetime.now().isoformat(timespec="seconds"),
        issue_type=issue_type,
        firm=normalise_space(match.firm).lower(),
        employee=match.employee,
        processed_date=match.processed_date_excel,
        indirect_request_id=match.indirect_request_id,
        details=details,
    )
    issues.append(issue)

    lines = [
        f"[{issue.recorded_at}] {issue.issue_type}",
        f"Firm: {issue.firm}",
        f"Employee: {issue.employee}",
        f"Processed Date: {issue.processed_date}",
        f"Request ID: {issue.indirect_request_id or 'N/A'}",
        f"Details: {issue.details}",
        "",
    ]
    with issue_log_path.open("a", encoding="utf-8") as handle:
        handle.write("\n".join(lines))

    log.warning(
        "Phase 2 issue recorded (%s) for %s / %s: %s",
        issue_type,
        issue.employee,
        issue.processed_date,
        details,
    )


def resolve_unique_target_path(
    requested_target: Path,
    issue_log_path: Path,
    issues: list[Phase2Issue],
    match: MatchRecord,
) -> Path:
    if not requested_target.exists():
        return requested_target

    counter = 2
    while True:
        candidate = requested_target.with_name(
            f"{requested_target.stem} (DuplicateName-{counter}){requested_target.suffix}"
        )
        if not candidate.exists():
            append_phase2_issue(
                issue_log_path=issue_log_path,
                issues=issues,
                match=match,
                issue_type="naming_conflict",
                details=(
                    f"Target file '{requested_target.name}' already existed. "
                    f"Saved this download as '{candidate.name}' instead."
                ),
            )
            return candidate
        counter += 1


def wait_for(
    driver: webdriver.Chrome,
    condition,
    timeout: int = WAIT_TIMEOUT,
):
    return WebDriverWait(driver, timeout).until(condition)


def wait_visible(driver: webdriver.Chrome, by: By, value: str, timeout: int = WAIT_TIMEOUT):
    return wait_for(driver, EC.visibility_of_element_located((by, value)), timeout)


def wait_clickable(driver: webdriver.Chrome, by: By, value: str, timeout: int = WAIT_TIMEOUT):
    return wait_for(driver, EC.element_to_be_clickable((by, value)), timeout)


def save_debug_artifacts(driver: webdriver.Chrome, artifacts_dir: Path, label: str) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    png_path = artifacts_dir / f"{timestamp}_{label}.png"
    html_path = artifacts_dir / f"{timestamp}_{label}.html"

    try:
        driver.save_screenshot(str(png_path))
        html_path.write_text(driver.page_source, encoding="utf-8")
        log.info("Saved debug artifacts: %s", png_path)
    except Exception as exc:
        log.warning("Could not save debug artifacts for %s: %s", label, exc)


def snapshot_download_dir(download_dir: str | Path) -> set[str]:
    return set(os.listdir(download_dir))


def wait_for_new_download(
    download_dir: str | Path,
    before_files: set[str],
    timeout: int = DOWNLOAD_TIMEOUT,
) -> str | None:
    deadline = time.time() + timeout
    download_dir = str(download_dir)

    while time.time() < deadline:
        current_files = set(os.listdir(download_dir))
        incomplete = {name for name in current_files if name.endswith(".crdownload") or name.endswith(".tmp")}
        completed = current_files - incomplete
        new_files = completed - before_files
        if new_files:
            newest = max((Path(download_dir) / name for name in new_files), key=lambda p: p.stat().st_mtime)
            return str(newest)
        time.sleep(1)
    return None


def scroll_into_view(driver: webdriver.Chrome, element: WebElement) -> None:
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
        element,
    )


def safe_click(driver: webdriver.Chrome, element: WebElement, description: str) -> None:
    scroll_into_view(driver, element)
    try:
        element.click()
        return
    except ElementClickInterceptedException:
        pass
    except StaleElementReferenceException:
        raise

    driver.execute_script("arguments[0].click();", element)
    log.info("Used JavaScript click for %s", description)


def first_visible(driver: webdriver.Chrome, selectors: Iterable[tuple[By, str]], timeout: int = 10) -> WebElement:
    end_time = datetime.now().timestamp() + timeout
    last_error = None

    while datetime.now().timestamp() < end_time:
        for by, value in selectors:
            try:
                element = driver.find_element(by, value)
                if element.is_displayed():
                    return element
            except Exception as exc:
                last_error = exc
        WebDriverWait(driver, 1).until(lambda _: True)

    raise TimeoutException(f"Could not find visible element. Last error: {last_error}")


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_manifest(excel_path: str) -> list[ExcelRecord]:
    log.info("Reading Excel manifest: %s", excel_path)

    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    worksheet = workbook.active
    rows = list(worksheet.iter_rows(values_only=True))

    if not rows:
        raise ValueError("Excel file is empty.")

    headers = [normalise_space(str(cell or "")).casefold() for cell in rows[0]]

    def find_column(name: str) -> int:
        key = name.casefold()
        for index, header in enumerate(headers):
            if key in header:
                return index
        raise KeyError(f"Column containing '{name}' was not found in Excel headers: {headers}")

    idx_firm = find_column("firm")
    idx_employee = find_column("employee")
    idx_processed_date = find_column("processed date")
    idx_title = find_column("title of piece")
    idx_status = find_column("status")

    records: list[ExcelRecord] = []

    for row in rows[1:]:
        status = normalise_space(str(row[idx_status] or ""))
        if status.casefold() != "approved":
            continue

        employee = normalise_space(str(row[idx_employee] or ""))
        processed_date = parse_date_value(row[idx_processed_date])
        if not employee or processed_date is None:
            continue

        records.append(
            ExcelRecord(
                firm=normalise_space(str(row[idx_firm] or "")),
                employee=employee,
                processed_date=processed_date,
                title=normalise_space(str(row[idx_title] or "")),
                status=status,
            )
        )

    workbook.close()
    log.info("Found %d approved records in Excel.", len(records))
    return records


# ============================================================================
# DRIVER
# ============================================================================

def create_driver(download_dir: str) -> webdriver.Chrome:
    ensure_dir(download_dir)

    options = ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-save-password-bubble")
    options.add_argument("--disable-features=PasswordManagerOnboarding,AutofillServerCommunication")

    prefs = {
        "download.default_directory": str(Path(download_dir).resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(1)
    return driver


# ============================================================================
# LOGIN + NAVIGATION
# ============================================================================

def login(driver: webdriver.Chrome, artifacts_dir: Path) -> None:
    log.info("Opening target URL: %s", REQUESTS_URL)
    driver.get(REQUESTS_URL)

    try:
        wait_for(lambda d: d.execute_script("return document.readyState") == "complete")
    except Exception:
        pass

    current_url = driver.current_url.lower()
    if "login" not in current_url and "membership" not in current_url:
        log.info("Already authenticated. Current URL: %s", driver.current_url)
        return

    if not PASSWORD.strip():
        raise ValueError("PASSWORD is blank. Enter your password in the PASSWORD variable.")

    log.info("Login page detected. Filling credentials.")

    email_selectors = [
        (By.CSS_SELECTOR, "input[type='email']"),
        (By.CSS_SELECTOR, "input[name*='email' i]"),
        (By.CSS_SELECTOR, "input[id*='email' i]"),
        (By.CSS_SELECTOR, "input[name*='user' i]"),
        (By.CSS_SELECTOR, "input[id*='user' i]"),
    ]
    password_selectors = [
        (By.CSS_SELECTOR, "input[type='password']"),
    ]

    try:
        email_field = first_visible(driver, email_selectors, timeout=20)
        password_field = first_visible(driver, password_selectors, timeout=20)

        email_field.clear()
        email_field.send_keys(EMAIL)
        password_field.clear()
        password_field.send_keys(PASSWORD)

        submit_candidates = [
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "input[type='submit']"),
            (By.XPATH, "//button[contains(., 'Log in') or contains(., 'Login') or contains(., 'Sign in')]"),
            (By.XPATH, "//input[@value='Log in' or @value='Login' or @value='Sign in']"),
        ]

        clicked = False
        for by, selector in submit_candidates:
            try:
                submit_button = wait_clickable(driver, by, selector, timeout=5)
                safe_click(driver, submit_button, "login submit button")
                clicked = True
                break
            except TimeoutException:
                continue

        if not clicked:
            password_field.send_keys(Keys.ENTER)

        wait_for(
            driver,
            lambda d: "login" not in d.current_url.lower() and "membership/login" not in d.current_url.lower(),
            timeout=45,
        )
        log.info("Login successful. Current URL: %s", driver.current_url)
    except Exception:
        save_debug_artifacts(driver, artifacts_dir, "login_failure")
        raise


def open_marketing_requests(driver: webdriver.Chrome, artifacts_dir: Path) -> None:
    log.info("Opening Marketing Material Requests page.")
    driver.get(REQUESTS_URL)
    wait_for(driver, lambda d: d.execute_script("return document.readyState") == "complete")

    if "marketingapproverpreclearancerequests" in driver.current_url.lower():
        log.info("Reached requests page directly.")
        return

    log.info("Direct navigation did not land on the requests grid. Falling back to dashboard card.")
    driver.get(DASHBOARD_URL)

    try:
        card = wait_clickable(
            driver,
            By.XPATH,
            "//a[contains(., 'Marketing Material Requests')] | //div[contains(., 'Marketing Material Requests')]",
            timeout=20,
        )
        safe_click(driver, card, "Marketing Material Requests dashboard card")
        wait_for(
            driver,
            lambda d: "marketingapproverpreclearancerequests" in d.current_url.lower(),
            timeout=30,
        )
        log.info("Reached requests page from dashboard.")
    except Exception:
        save_debug_artifacts(driver, artifacts_dir, "navigation_failure")
        raise


def switch_to_processed(driver: webdriver.Chrome, artifacts_dir: Path) -> None:
    log.info("Switching grid to Processed.")
    try:
        wait_visible(
            driver,
            By.XPATH,
            "//*[contains(normalize-space(.), 'Marketing Material Requests')]",
            timeout=20,
        )

        processed_selectors = [
            (By.XPATH, "//button[normalize-space()='Processed']"),
            (By.XPATH, "//a[normalize-space()='Processed']"),
            (
                By.XPATH,
                "//*[self::div or self::span or self::li][normalize-space()='Processed' and "
                "ancestor::*[contains(@class, 'k-tabstrip') or contains(@class, 'k-widget') or contains(@class, 'k-grid') or contains(@class, 'btn-group')]]",
            ),
            (
                By.XPATH,
                "//*[normalize-space(text())='Processed' and "
                "ancestor::*[contains(@class, 'k-grid') or contains(@class, 'k-widget') or contains(@class, 'tab') or contains(@class, 'btn-group')]]",
            ),
        ]

        processed_tab = first_visible(driver, processed_selectors, timeout=20)
        safe_click(driver, processed_tab, "Processed tab")

        wait_for(
            driver,
            lambda d: len(d.find_elements(By.XPATH, "//table//th[normalize-space()='Username']")) > 0,
            timeout=20,
        )
        log.info("Processed tab activated.")
    except Exception:
        save_debug_artifacts(driver, artifacts_dir, "processed_tab_failure")
        raise


def build_processed_date_filter_groups(excel_records: list[ExcelRecord]) -> list[dict]:
    unique_dates = sorted({record.processed_date for record in excel_records})
    groups: list[dict] = []

    for processed_day in unique_dates:
        next_day = processed_day + timedelta(days=1)
        groups.append(
            {
                "logic": "and",
                "filters": [
                    {"field": "ApprovedDate", "operator": "gte", "value": processed_day.isoformat()},
                    {"field": "ApprovedDate", "operator": "lt", "value": next_day.isoformat()},
                ],
            }
        )

    return groups


def apply_initial_keyword_filter(
    driver: webdriver.Chrome,
    keyword: str,
    artifacts_dir: Path,
) -> None:
    log.info("Applying initial processed-grid filter for keyword '%s'.", keyword)

    try:
        wait_visible(driver, By.ID, "mm-processed", timeout=20)

        script_result = driver.execute_async_script(
            """
            const done = arguments[arguments.length - 1];
            const keyword = arguments[0];
            const gridElement = document.getElementById('mm-processed');
            if (!gridElement || !window.jQuery) {
                done({ ok: false, reason: 'grid_or_jquery_missing' });
                return;
            }

            const grid = window.jQuery(gridElement).data('kendoGrid');
            if (!grid) {
                done({ ok: false, reason: 'kendo_grid_missing' });
                return;
            }

            const finalFilter = {
                logic: 'and',
                filters: [
                    { field: 'UserName', operator: 'contains', value: keyword },
                    { field: 'StatusName', operator: 'eq', value: 'Approved' }
                ]
            };

            grid.one('dataBound', function () {
                done({
                    ok: true,
                    total: grid.dataSource.total(),
                    viewLength: grid.dataSource.view().length,
                    pageSize: grid.dataSource.pageSize()
                });
            });

            grid.dataSource.page(1);
            grid.dataSource.filter(finalFilter);
            """,
            keyword,
        )

        if not script_result or not script_result.get("ok"):
            raise RuntimeError(f"Could not apply initial grid filter via grid API: {script_result}")

        log.info(
            "Initial grid filter applied. total=%s current_page_rows=%s page_size=%s",
            script_result.get("total"),
            script_result.get("viewLength"),
            script_result.get("pageSize"),
        )
    except Exception:
        save_debug_artifacts(driver, artifacts_dir, "filter_failure")
        raise


# ============================================================================
# GRID MATCHING
# ============================================================================

def group_excel_records(records: list[ExcelRecord]) -> list[ExcelRecord]:
    grouped: dict[tuple[str, date, str], ExcelRecord] = {}
    for record in records:
        key = (
            canonical_person_name(record.employee),
            record.processed_date,
            normalise_space(record.firm).casefold(),
        )
        grouped.setdefault(key, record)
    return list(grouped.values())


def fetch_web_records_for_excel_record(
    driver: webdriver.Chrome,
    excel_record: ExcelRecord,
    keyword: str,
) -> list[WebRecord]:
    next_day = excel_record.processed_date + timedelta(days=1)
    log.info(
        "Matching web rows for employee='%s' processed_date=%s",
        excel_record.employee,
        excel_record.processed_date.isoformat(),
    )

    page_items = driver.execute_async_script(
        """
        const done = arguments[arguments.length - 1];
        const keyword = arguments[0];
        const employee = arguments[1];
        const startDate = arguments[2];
        const endDate = arguments[3];

        const gridElement = document.getElementById('mm-processed');
        if (!gridElement || !window.jQuery) {
            done({ ok: false, reason: 'grid_or_jquery_missing' });
            return;
        }

        const grid = window.jQuery(gridElement).data('kendoGrid');
        if (!grid) {
            done({ ok: false, reason: 'kendo_grid_missing' });
            return;
        }

        const finalFilter = {
            logic: 'and',
            filters: [
                { field: 'UserName', operator: 'contains', value: keyword },
                { field: 'StatusName', operator: 'eq', value: 'Approved' },
                { field: 'EmployeeName', operator: 'eq', value: employee },
                { field: 'ApprovedDate', operator: 'gte', value: new Date(startDate) },
                { field: 'ApprovedDate', operator: 'lt', value: new Date(endDate) }
            ]
        };

        function serializeRows() {
            const view = grid.dataSource.view();
            const rows = [];
            for (let i = 0; i < view.length; i++) {
                const item = view[i];
                rows.push({
                    employee: item.EmployeeName || '',
                    username: item.UserName || '',
                    processed_date_text: item.ApprovedDateWithTimeZone || '',
                    status: item.StatusName || '',
                    title: item.TitleOfPiece || '',
                    indirect_request_id: item.IndirectRequestId || '',
                    row_index: i + 1
                });
            }
            return rows;
        }

        grid.one('dataBound', function () {
            done({
                ok: true,
                total: grid.dataSource.total(),
                rows: serializeRows()
            });
        });

        grid.dataSource.page(1);
        grid.dataSource.filter(finalFilter);
        """,
        keyword,
        excel_record.employee,
        excel_record.processed_date.isoformat(),
        next_day.isoformat(),
    )

    if not page_items or not page_items.get("ok"):
        raise RuntimeError(
            f"Could not fetch web rows for {excel_record.employee} | {excel_record.processed_date}: {page_items}"
        )

    rows = page_items.get("rows", [])
    total = int(page_items.get("total") or 0)
    if total > len(rows):
        log.warning(
            "Grid returned total=%d but only %d rows on first page for %s | %s. "
            "Keeping first-page rows only.",
            total,
            len(rows),
            excel_record.employee,
            excel_record.processed_date,
        )

    page_records: list[WebRecord] = []
    for item in rows:
        employee = normalise_space(item.get("employee", ""))
        processed_date_text = normalise_space(item.get("processed_date_text", ""))
        processed_date = parse_date_from_web_text(processed_date_text)
        if not employee or processed_date is None:
            continue

        indirect_request_id = item.get("indirect_request_id") or None
        view_href = (
            f"{BASE_URL}/Preclear/PreclearMarketingMaterialRequestDetails/{indirect_request_id}"
            f"?navigatedFromPage=MarketingApproverPreclearanceRequests&navigatedFromController=PreClear"
            if indirect_request_id
            else None
        )

        page_records.append(
            WebRecord(
                employee=employee,
                username=normalise_space(item.get("username", "")),
                processed_date=processed_date,
                processed_date_text=processed_date_text,
                status=normalise_space(item.get("status", "")),
                title=normalise_space(item.get("title", "")),
                row_page=1,
                row_index=int(item.get("row_index", 0) or 0),
                view_href=view_href,
                indirect_request_id=indirect_request_id,
            )
        )

    log.info(
        "Fetched %d candidate web row(s) for %s | %s",
        len(page_records),
        excel_record.employee,
        excel_record.processed_date.isoformat(),
    )
    return page_records


# ============================================================================
# MATCHING
# ============================================================================

def match_records(excel_records: list[ExcelRecord], web_records: list[WebRecord]) -> list[MatchRecord]:
    matched: list[MatchRecord] = []

    approved_web_records = [record for record in web_records if record.status.casefold() == "approved"]

    for excel_record in excel_records:
        for web_record in approved_web_records:
            if canonical_person_name(excel_record.employee) != canonical_person_name(web_record.employee):
                continue
            if excel_record.processed_date != web_record.processed_date:
                continue

            matched.append(
                MatchRecord(
                    firm=excel_record.firm,
                    employee=excel_record.employee,
                    processed_date_excel=excel_record.processed_date.isoformat(),
                    processed_date_web=web_record.processed_date_text,
                    username=web_record.username,
                    title=web_record.title,
                    status=web_record.status,
                    row_page=web_record.row_page,
                    row_index=web_record.row_index,
                    view_href=web_record.view_href,
                    indirect_request_id=web_record.indirect_request_id,
                )
            )

    log.info("Matched %d records.", len(matched))
    return matched


def unique_match_records(matches: list[MatchRecord]) -> list[MatchRecord]:
    unique: list[MatchRecord] = []
    seen: set[str] = set()

    for match in matches:
        key = match.indirect_request_id or f"{match.employee}|{match.processed_date_excel}|{match.view_href}"
        if key in seen:
            continue
        seen.add(key)
        unique.append(match)

    return unique


# ============================================================================
# PHASE 2 DOWNLOADS
# ============================================================================

def open_match_detail_page(driver: webdriver.Chrome, match: MatchRecord, artifacts_dir: Path) -> None:
    if not match.view_href:
        raise RuntimeError(f"No detail URL available for {match.employee} | {match.processed_date_excel}")

    log.info("Opening detail page for %s | %s", match.employee, match.processed_date_excel)
    driver.get(match.view_href)
    try:
        wait_visible(
            driver,
            By.XPATH,
            "//*[contains(normalize-space(.), 'Marketing Material Request Details')]",
            timeout=25,
        )
    except Exception:
        save_debug_artifacts(driver, artifacts_dir, "detail_page_failure")
        raise


def expand_detail_section(driver: webdriver.Chrome, section_text: str) -> None:
    candidates = driver.find_elements(
        By.XPATH,
        f"//*[normalize-space(text())='{section_text}' and (self::a or self::span or self::div or self::button)]",
    )
    if not candidates:
        log.warning("Section '%s' not found on detail page.", section_text)
        return

    element = candidates[0]
    try:
        safe_click(driver, element, section_text)
        time.sleep(1)
    except Exception as exc:
        log.warning("Could not expand '%s': %s", section_text, exc)


def expand_detail_sections(driver: webdriver.Chrome) -> None:
    expand_detail_section(driver, "View Communications")
    expand_detail_section(driver, "View Supervisor Notes")


def collect_attach_document_links(driver: webdriver.Chrome) -> list[WebElement]:
    link_candidates = driver.find_elements(
        By.XPATH,
        "//div[contains(@class, 'main-attachment-content')]//a[contains(@class, 'doc-link') or contains(@class, 'doc-name') or @download]",
    )

    anchor_candidates: list[WebElement] = []
    seen: set[str] = set()
    for link in link_candidates:
        href = (link.get_attribute("href") or "").strip()
        text = normalise_space(link.text)
        if not href or href.lower().startswith("javascript:"):
            continue
        if "/Document/DownloadFile" not in href and not link.get_attribute("download"):
            continue
        key = f"{href}|{text}"
        if key in seen:
            continue
        seen.add(key)
        anchor_candidates.append(link)

    return anchor_candidates


def download_attached_documents(
    driver: webdriver.Chrome,
    download_dir: str | Path,
) -> list[str]:
    downloaded_files: list[str] = []
    doc_links = collect_attach_document_links(driver)
    if not doc_links:
        log.info("No attached documents found on this detail page.")
        return downloaded_files

    log.info("Found %d attached document link(s).", len(doc_links))
    for index, link in enumerate(doc_links, start=1):
        label = normalise_space(link.text) or f"doc_{index}"
        before_files = snapshot_download_dir(download_dir)
        safe_click(driver, link, f"attached document link {index}")
        new_file = wait_for_new_download(download_dir, before_files)
        if not new_file:
            raise RuntimeError(f"Timed out downloading attached document: {label}")
        log.info("Downloaded attached document %d: %s", index, new_file)
        downloaded_files.append(new_file)
        time.sleep(1)

    return downloaded_files


def download_approval_pdf(
    driver: webdriver.Chrome,
    download_dir: str | Path,
) -> str | None:
    buttons = driver.find_elements(
        By.XPATH,
        "//button[contains(normalize-space(.), 'Export to PDF')] | //a[contains(normalize-space(.), 'Export to PDF')]",
    )
    if not buttons:
        log.warning("Export to PDF button not found.")
        return None

    before_files = snapshot_download_dir(download_dir)
    safe_click(driver, buttons[0], "Export to PDF")
    new_file = wait_for_new_download(download_dir, before_files)
    if not new_file:
        raise RuntimeError("Timed out downloading approval PDF.")
    log.info("Downloaded approval PDF: %s", new_file)
    return new_file


def rename_downloaded_files(
    match: MatchRecord,
    doc_files: list[str],
    approval_file: str | None,
    output_dir: Path,
    approval_counters: dict[tuple[str, str], int],
    doc_counters: dict[tuple[str, str], int],
    issue_log_path: Path,
    issues: list[Phase2Issue],
) -> DownloadOutcome:
    firm = normalise_space(match.firm).lower()
    processed_date = datetime.fromisoformat(match.processed_date_excel).date()
    date_str = processed_date.strftime("%m.%d.%Y")
    key = (firm, date_str)

    final_doc_files: list[str] = []
    for doc_file in doc_files:
        source = Path(doc_file)
        ext = source.suffix or ".docx"
        doc_counters[key] = doc_counters.get(key, 0) + 1
        serial = doc_counters[key]
        target = output_dir / f"{firm} - {date_str} - Doc - {serial}{ext}"
        target = resolve_unique_target_path(target, issue_log_path, issues, match)
        shutil.move(str(source), str(target))
        final_doc_files.append(str(target))

    final_approval_file: str | None = None
    if approval_file:
        approval_counters[key] = approval_counters.get(key, 0) + 1
        serial = approval_counters[key]
        suffix = " ( No Doc )" if not final_doc_files else ""
        target = output_dir / f"{firm} - {date_str} - Approval - {serial}{suffix}.pdf"
        target = resolve_unique_target_path(target, issue_log_path, issues, match)
        shutil.move(str(approval_file), str(target))
        final_approval_file = str(target)

    return DownloadOutcome(
        firm=firm,
        employee=match.employee,
        processed_date=match.processed_date_excel,
        indirect_request_id=match.indirect_request_id,
        approval_file=final_approval_file,
        doc_files=final_doc_files,
    )


def process_phase_2_downloads(
    driver: webdriver.Chrome,
    matches: list[MatchRecord],
    output_dir: Path,
    artifacts_dir: Path,
    download_dir: str | Path,
) -> tuple[list[DownloadOutcome], list[Phase2Issue], Path]:
    unique_matches = unique_match_records(matches)
    log.info("Phase 2 will process %d unique matched record(s).", len(unique_matches))

    approval_counters: dict[tuple[str, str], int] = {}
    doc_counters: dict[tuple[str, str], int] = {}
    outcomes: list[DownloadOutcome] = []
    issues: list[Phase2Issue] = []
    issue_log_path = output_dir / DOWNLOAD_ISSUES_NAME

    issue_log_path.write_text(
        "Phase 2 download issue log\n"
        f"Generated at: {datetime.now().isoformat(timespec='seconds')}\n"
        "This file records download failures, naming conflicts, and other recoverable issues.\n\n",
        encoding="utf-8",
    )

    for index, match in enumerate(unique_matches, start=1):
        log.info("Phase 2 record %d/%d", index, len(unique_matches))
        try:
            open_match_detail_page(driver, match, artifacts_dir)
            expand_detail_sections(driver)
            doc_files = download_attached_documents(driver, download_dir)
            approval_file = download_approval_pdf(driver, download_dir)
            if not approval_file:
                append_phase2_issue(
                    issue_log_path=issue_log_path,
                    issues=issues,
                    match=match,
                    issue_type="approval_pdf_missing",
                    details="Export to PDF button was not found or no approval PDF was downloaded for this record.",
                )
            outcome = rename_downloaded_files(
                match=match,
                doc_files=doc_files,
                approval_file=approval_file,
                output_dir=output_dir,
                approval_counters=approval_counters,
                doc_counters=doc_counters,
                issue_log_path=issue_log_path,
                issues=issues,
            )
            outcomes.append(outcome)
        except Exception as exc:
            save_debug_artifacts(driver, artifacts_dir, f"phase2_record_{index}_failure")
            append_phase2_issue(
                issue_log_path=issue_log_path,
                issues=issues,
                match=match,
                issue_type="download_failure",
                details=str(exc),
            )
            continue

    return outcomes, issues, issue_log_path


def save_phase2_summary(output_dir: Path, outcomes: list[DownloadOutcome]) -> Path:
    output_path = output_dir / DOWNLOAD_SUMMARY_NAME
    payload = {
        "generated_at": datetime.now().isoformat(),
        "download_count": len(outcomes),
        "records": [asdict(outcome) for outcome in outcomes],
    }
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return output_path


def save_phase1_output(output_dir: Path, matches: list[MatchRecord], web_records: list[WebRecord]) -> Path:
    output_path = output_dir / JSON_OUTPUT_NAME

    payload = {
        "generated_at": datetime.now().isoformat(),
        "match_count": len(matches),
        "web_record_count": len(web_records),
        "matches": [asdict(match) for match in matches],
    }

    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return output_path


# ============================================================================
# MAIN
# ============================================================================

def main() -> None:
    log.info("=" * 72)
    log.info("ComplySci Marketing Material Downloader - PHASE 1 START")
    log.info("=" * 72)

    missing_config: list[str] = []
    if not EMAIL:
        missing_config.append("EMAIL")
    if not PASSWORD.strip():
        missing_config.append("PASSWORD")
    if not FILTER_KEYWORD:
        missing_config.append("FILTER_KEYWORD")
    if not EXCEL_PATH:
        missing_config.append("EXCEL_PATH")
    if not OUTPUT_FOLDER:
        missing_config.append("OUTPUT_FOLDER")

    if missing_config:
        log.error(
            "Missing required configuration in %s: %s",
            Path(__file__).with_name(ENV_FILE_NAME),
            ", ".join(missing_config),
        )
        return

    output_dir = ensure_dir(OUTPUT_FOLDER)
    artifacts_dir = ensure_dir(output_dir / ARTIFACTS_DIRNAME)
    temp_download_dir = ensure_dir(output_dir / "_temp_downloads_phase1")

    driver: webdriver.Chrome | None = None

    try:
        excel_records = read_excel_manifest(EXCEL_PATH)
        if not excel_records:
            log.error("No approved rows were found in the Excel file.")
            return

        driver = create_driver(str(temp_download_dir))

        login(driver, artifacts_dir)
        open_marketing_requests(driver, artifacts_dir)
        switch_to_processed(driver, artifacts_dir)
        apply_initial_keyword_filter(driver, FILTER_KEYWORD, artifacts_dir)

        web_records: list[WebRecord] = []
        grouped_excel_records = group_excel_records(excel_records)
        log.info("Matching against %d unique approved Excel employee/date combination(s).", len(grouped_excel_records))
        for excel_record in grouped_excel_records:
            web_records.extend(fetch_web_records_for_excel_record(driver, excel_record, FILTER_KEYWORD))

        if not web_records:
            save_debug_artifacts(driver, artifacts_dir, "no_web_rows")
            log.error("No rows were scraped from the website after filtering.")
            return

        matches = match_records(excel_records, web_records)
        output_path = save_phase1_output(output_dir, matches, web_records)

        if not matches:
            save_debug_artifacts(driver, artifacts_dir, "no_matches")
            log.warning("Phase 1 ran successfully, but no Excel rows matched the website rows.")
        else:
            log.info("Phase 1 completed successfully.")

        log.info("Phase 1 JSON output: %s", output_path)

        if matches:
            phase2_outcomes, phase2_issues, phase2_issue_log = process_phase_2_downloads(
                driver=driver,
                matches=matches,
                output_dir=output_dir,
                artifacts_dir=artifacts_dir,
                download_dir=temp_download_dir,
            )
            phase2_summary = save_phase2_summary(output_dir, phase2_outcomes)
            log.info("Phase 2 completed successfully. Download summary: %s", phase2_summary)
            if phase2_issues:
                log.warning(
                    "Phase 2 completed with %d logged issue(s). Review: %s",
                    len(phase2_issues),
                    phase2_issue_log,
                )
            else:
                log.info("Phase 2 completed with no logged download/naming issues.")

    except Exception as exc:
        log.critical("Fatal error: %s", exc)
        log.debug(traceback.format_exc())
        if driver is not None:
            save_debug_artifacts(driver, artifacts_dir, "fatal_error")
        raise
    finally:
        try:
            remaining_files = list(Path(temp_download_dir).iterdir())
            if not remaining_files:
                Path(temp_download_dir).rmdir()
        except Exception:
            pass
        if driver is not None:
            driver.quit()
            log.info("Browser closed.")


if __name__ == "__main__":
    main()
