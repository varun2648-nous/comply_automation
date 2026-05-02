"""
Phase 1 - ComplySci Marketing Material Request matcher

What this script does:
1. Opens ComplySci and logs in.
2. Navigates to Marketing Material Requests.
3. Switches the grid to Processed.
4. Applies the Username filter using a firm keyword.
5. Applies the Status filter so only Approved records remain.
6. Scrapes the filtered approved rows directly from the website.
7. Saves the phase-1 output to JSON so phase 2 can download from those rows.

What this script also does:
- download attached documents
- click Export to PDF
- rename files and log recoverable issues
"""

from __future__ import annotations

import json
import logging
import os
import math
import re
import shutil
import time
import traceback
from dataclasses import dataclass, asdict
from datetime import date, datetime
from pathlib import Path
from typing import Iterable
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
FIRM_NAME = os.getenv("FIRM_NAME", "").strip()
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "").strip()
RESUME_FROM_PROCESSED_DATE = os.getenv("RESUME_FROM_PROCESSED_DATE", "").strip()


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
RUN_SUMMARY_NAME = "phase2_run_summary.json"
DOWNLOAD_TIMEOUT = 45
GRID_FETCH_BATCH_SIZE = 1000
SCRIPT_TIMEOUT = 180


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


@dataclass
class RunSummary:
    generated_at: str
    keyword_filter: str
    firm_name: str
    record_count: int
    successful_record_count: int
    issue_count: int
    total_downloaded_files: int
    approval_pdf_count: int
    attached_document_count: int
    pdf_file_count: int
    word_file_count: int
    other_file_count: int
    elapsed_seconds: float


@dataclass
class NamedFileEntry:
    path: str
    suffix_text: str
    extension: str


# ============================================================================
# GENERIC HELPERS
# ============================================================================

def normalise_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def parse_date_from_web_text(raw_text: str) -> date | None:
    match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", raw_text or "")
    if not match:
        return None
    return datetime.strptime(match.group(1), "%m/%d/%Y").date()


def parse_resume_date_config(raw_text: str) -> date:
    value = normalise_space(raw_text)
    formats = ("%Y-%m-%d", "%m.%d.%Y", "%m/%d/%Y")

    for date_format in formats:
        try:
            return datetime.strptime(value, date_format).date()
        except ValueError:
            continue

    raise ValueError(
        "RESUME_FROM_PROCESSED_DATE must use YYYY-MM-DD, MM.DD.YYYY, or MM/DD/YYYY format."
    )


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


def build_download_target_path(
    output_dir: Path,
    firm: str,
    date_str: str,
    label: str,
    extension: str,
    serial: int | None = None,
    suffix_text: str = "",
) -> Path:
    serial_text = f" - {serial}" if serial is not None else ""
    return output_dir / f"{firm} - {date_str} - {label}{serial_text}{suffix_text}{extension}"


def move_download_to_target(
    source: Path,
    requested_target: Path,
    issue_log_path: Path,
    issues: list[Phase2Issue],
    match: MatchRecord,
) -> Path:
    if source.resolve() == requested_target.resolve():
        return requested_target

    target = resolve_unique_target_path(requested_target, issue_log_path, issues, match)
    shutil.move(str(source), str(target))
    return target


def assign_named_download(
    source_file: str,
    output_dir: Path,
    firm: str,
    date_str: str,
    label: str,
    extension: str,
    suffix_text: str,
    naming_registry: dict[tuple[str, str, str], list[NamedFileEntry]],
    issue_log_path: Path,
    issues: list[Phase2Issue],
    match: MatchRecord,
) -> str:
    registry_key = (firm, date_str, label)
    existing_entries = naming_registry.setdefault(registry_key, [])
    source_path = Path(source_file)

    if not existing_entries:
        target = build_download_target_path(
            output_dir=output_dir,
            firm=firm,
            date_str=date_str,
            label=label,
            extension=extension,
            suffix_text=suffix_text,
        )
        final_target = move_download_to_target(source_path, target, issue_log_path, issues, match)
        existing_entries.append(
            NamedFileEntry(
                path=str(final_target),
                suffix_text=suffix_text,
                extension=extension,
            )
        )
        return str(final_target)

    if len(existing_entries) == 1:
        first_entry = existing_entries[0]
        first_source = Path(first_entry.path)
        first_target = build_download_target_path(
            output_dir=output_dir,
            firm=firm,
            date_str=date_str,
            label=label,
            extension=first_entry.extension,
            serial=1,
            suffix_text=first_entry.suffix_text,
        )
        first_final_target = move_download_to_target(first_source, first_target, issue_log_path, issues, match)
        first_entry.path = str(first_final_target)

    serial = len(existing_entries) + 1
    target = build_download_target_path(
        output_dir=output_dir,
        firm=firm,
        date_str=date_str,
        label=label,
        extension=extension,
        serial=serial,
        suffix_text=suffix_text,
    )
    final_target = move_download_to_target(source_path, target, issue_log_path, issues, match)
    existing_entries.append(
        NamedFileEntry(
            path=str(final_target),
            suffix_text=suffix_text,
            extension=extension,
        )
    )
    return str(final_target)


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
    driver.set_script_timeout(SCRIPT_TIMEOUT)
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
# GRID RECORD COLLECTION
# ============================================================================

def fetch_filtered_grid_page(driver: webdriver.Chrome, page_number: int, page_size: int) -> dict:
    page_items = driver.execute_async_script(
        """
        const done = arguments[arguments.length - 1];
        const pageNumber = arguments[0];
        const requestedPageSize = arguments[1];

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

        function complete() {
            const activePageSize = grid.dataSource.pageSize() || requestedPageSize || 0;
            done({
                ok: true,
                total: grid.dataSource.total(),
                page: grid.dataSource.page() || pageNumber,
                pageSize: activePageSize,
                rows: serializeRows()
            });
        }

        const currentPage = grid.dataSource.page() || 1;
        const currentPageSize = grid.dataSource.pageSize() || 0;
        if (currentPage === pageNumber && currentPageSize === requestedPageSize) {
            window.setTimeout(complete, 0);
            return;
        }

        grid.one('dataBound', complete);

        grid.dataSource.query({
            page: pageNumber,
            pageSize: requestedPageSize,
            filter: grid.dataSource.filter(),
            sort: grid.dataSource.sort(),
            group: grid.dataSource.group()
        });
        """,
        page_number,
        page_size,
    )

    if not page_items or not page_items.get("ok"):
        raise RuntimeError(f"Could not fetch grid page {page_number} with page_size {page_size}: {page_items}")

    return page_items


def fetch_all_filtered_web_records(driver: webdriver.Chrome, keyword: str) -> list[WebRecord]:
    log.info("Collecting filtered approved records directly from the processed grid.")

    first_page = fetch_filtered_grid_page(driver, 1, 100)
    total = int(first_page.get("total") or 0)
    initial_page_size = int(first_page.get("pageSize") or 0)
    batch_page_size = min(max(initial_page_size, GRID_FETCH_BATCH_SIZE), total) if total else initial_page_size

    if total > 0 and batch_page_size and batch_page_size != initial_page_size:
        log.info(
            "Expanding grid fetch batch size from %d to %d to reduce pre-download paging.",
            initial_page_size,
            batch_page_size,
        )
        first_page = fetch_filtered_grid_page(driver, 1, batch_page_size)

    page_size = int(first_page.get("pageSize") or 0)
    if batch_page_size and page_size and page_size < batch_page_size:
        log.warning(
            "Grid kept page size at %d even after requesting %d. Collection will continue in %d page(s).",
            page_size,
            batch_page_size,
            max(1, math.ceil(total / page_size)),
        )
    total_pages = max(1, math.ceil(total / page_size)) if page_size else 1

    all_rows = []
    for row in first_page.get("rows", []):
        row["row_page"] = int(first_page.get("page") or 1)
        all_rows.append(row)
    log.info(
        "Filtered grid has total=%d row(s), page_size=%d, total_pages=%d",
        total,
        page_size,
        total_pages,
    )

    for page_number in range(2, total_pages + 1):
        page_items = fetch_filtered_grid_page(driver, page_number, page_size)
        for row in page_items.get("rows", []):
            row["row_page"] = int(page_items.get("page") or page_number)
            all_rows.append(row)

    page_records: list[WebRecord] = []
    for item in all_rows:
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
                row_page=int(item.get("row_page", 1) or 1),
                row_index=int(item.get("row_index", 0) or 0),
                view_href=view_href,
                indirect_request_id=indirect_request_id,
            )
        )

    log.info(
        "Collected %d approved web record(s) for keyword '%s'.",
        len(page_records),
        keyword,
    )
    return page_records


def build_matches_from_web_records(web_records: list[WebRecord], firm_label: str) -> list[MatchRecord]:
    matches: list[MatchRecord] = []

    for web_record in web_records:
        if web_record.status.casefold() != "approved":
            continue

        matches.append(
            MatchRecord(
                firm=firm_label,
                employee=web_record.employee,
                processed_date_excel=web_record.processed_date.isoformat(),
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

    log.info("Prepared %d phase-2 record(s) directly from web results.", len(matches))
    return matches


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


def load_completed_request_ids(output_dir: Path) -> set[str]:
    summary_path = output_dir / DOWNLOAD_SUMMARY_NAME
    if not summary_path.exists():
        log.info("No previous phase-2 summary found at %s; resume will not skip completed request IDs.", summary_path)
        return set()

    try:
        payload = json.loads(summary_path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning("Could not read previous phase-2 summary at %s: %s", summary_path, exc)
        return set()

    completed_ids: set[str] = set()
    for record in payload.get("records", []):
        request_id = normalise_space(record.get("indirect_request_id", ""))
        if request_id:
            completed_ids.add(request_id)

    log.info("Loaded %d completed request ID(s) from previous phase-2 summary.", len(completed_ids))
    return completed_ids


def filter_matches_for_resume(matches: list[MatchRecord], output_dir: Path) -> list[MatchRecord]:
    if not RESUME_FROM_PROCESSED_DATE:
        return matches

    resume_date = parse_resume_date_config(RESUME_FROM_PROCESSED_DATE)
    completed_ids = load_completed_request_ids(output_dir)
    filtered_matches: list[MatchRecord] = []
    skipped_newer = 0
    skipped_completed = 0

    for match in matches:
        processed_date = datetime.fromisoformat(match.processed_date_excel).date()
        if processed_date > resume_date:
            skipped_newer += 1
            continue

        if match.indirect_request_id and match.indirect_request_id in completed_ids:
            skipped_completed += 1
            continue

        filtered_matches.append(match)

    log.info(
        "Resume from processed date %s: kept %d record(s), skipped %d newer record(s), skipped %d already completed record(s).",
        resume_date.isoformat(),
        len(filtered_matches),
        skipped_newer,
        skipped_completed,
    )
    return filtered_matches


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
    naming_registry: dict[tuple[str, str, str], list[NamedFileEntry]],
    issue_log_path: Path,
    issues: list[Phase2Issue],
) -> DownloadOutcome:
    firm = normalise_space(match.firm).lower()
    processed_date = datetime.fromisoformat(match.processed_date_excel).date()
    date_str = processed_date.strftime("%m.%d.%Y")

    final_doc_files: list[str] = []
    for doc_file in doc_files:
        source = Path(doc_file)
        ext = source.suffix or ".docx"
        final_target = assign_named_download(
            source_file=str(source),
            output_dir=output_dir,
            firm=firm,
            date_str=date_str,
            label="Doc",
            extension=ext,
            suffix_text="",
            naming_registry=naming_registry,
            issue_log_path=issue_log_path,
            issues=issues,
            match=match,
        )
        final_doc_files.append(final_target)

    final_approval_file: str | None = None
    if approval_file:
        suffix = " ( No Doc )" if not final_doc_files else ""
        final_approval_file = assign_named_download(
            source_file=str(approval_file),
            output_dir=output_dir,
            firm=firm,
            date_str=date_str,
            label="Approval",
            extension=".pdf",
            suffix_text=suffix,
            naming_registry=naming_registry,
            issue_log_path=issue_log_path,
            issues=issues,
            match=match,
        )

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

    naming_registry: dict[tuple[str, str, str], list[NamedFileEntry]] = {}
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
                naming_registry=naming_registry,
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


def build_run_summary(
    keyword_filter: str,
    firm_name: str,
    record_count: int,
    outcomes: list[DownloadOutcome],
    issues: list[Phase2Issue],
    elapsed_seconds: float,
) -> RunSummary:
    total_downloaded_files = 0
    approval_pdf_count = 0
    attached_document_count = 0
    pdf_file_count = 0
    word_file_count = 0
    other_file_count = 0
    word_extensions = {".doc", ".docx"}

    for outcome in outcomes:
        if outcome.approval_file:
            approval_pdf_count += 1
            total_downloaded_files += 1
            pdf_file_count += 1

        attached_document_count += len(outcome.doc_files)
        total_downloaded_files += len(outcome.doc_files)

        for file_path in outcome.doc_files:
            suffix = Path(file_path).suffix.casefold()
            if suffix == ".pdf":
                pdf_file_count += 1
            elif suffix in word_extensions:
                word_file_count += 1
            else:
                other_file_count += 1

    return RunSummary(
        generated_at=datetime.now().isoformat(),
        keyword_filter=keyword_filter,
        firm_name=firm_name,
        record_count=record_count,
        successful_record_count=len(outcomes),
        issue_count=len(issues),
        total_downloaded_files=total_downloaded_files,
        approval_pdf_count=approval_pdf_count,
        attached_document_count=attached_document_count,
        pdf_file_count=pdf_file_count,
        word_file_count=word_file_count,
        other_file_count=other_file_count,
        elapsed_seconds=round(elapsed_seconds, 2),
    )


def save_run_summary(output_dir: Path, summary: RunSummary) -> Path:
    output_path = output_dir / RUN_SUMMARY_NAME
    output_path.write_text(json.dumps(asdict(summary), indent=2), encoding="utf-8")
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
    if not FIRM_NAME:
        missing_config.append("FIRM_NAME")
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
    run_started_at = time.time()

    driver: webdriver.Chrome | None = None

    try:
        driver = create_driver(str(temp_download_dir))

        login(driver, artifacts_dir)
        open_marketing_requests(driver, artifacts_dir)
        switch_to_processed(driver, artifacts_dir)
        apply_initial_keyword_filter(driver, FILTER_KEYWORD, artifacts_dir)

        web_records = fetch_all_filtered_web_records(driver, FILTER_KEYWORD)

        if not web_records:
            save_debug_artifacts(driver, artifacts_dir, "no_web_rows")
            run_summary = build_run_summary(
                keyword_filter=FILTER_KEYWORD,
                firm_name=FIRM_NAME,
                record_count=0,
                outcomes=[],
                issues=[],
                elapsed_seconds=time.time() - run_started_at,
            )
            save_run_summary(output_dir, run_summary)
            log.error("No rows were scraped from the website after filtering.")
            return

        matches = build_matches_from_web_records(web_records, FIRM_NAME)
        output_path = save_phase1_output(output_dir, matches, web_records)

        if not matches:
            save_debug_artifacts(driver, artifacts_dir, "no_matches")
            run_summary = build_run_summary(
                keyword_filter=FILTER_KEYWORD,
                firm_name=FIRM_NAME,
                record_count=0,
                outcomes=[],
                issues=[],
                elapsed_seconds=time.time() - run_started_at,
            )
            save_run_summary(output_dir, run_summary)
            log.warning("Phase 1 ran successfully, but no approved website rows were prepared for download.")
        else:
            log.info("Phase 1 completed successfully.")

        log.info("Phase 1 JSON output: %s", output_path)

        phase2_matches = filter_matches_for_resume(matches, output_dir)

        if phase2_matches:
            phase2_outcomes, phase2_issues, phase2_issue_log = process_phase_2_downloads(
                driver=driver,
                matches=phase2_matches,
                output_dir=output_dir,
                artifacts_dir=artifacts_dir,
                download_dir=temp_download_dir,
            )
            phase2_summary = save_phase2_summary(output_dir, phase2_outcomes)
            log.info("Phase 2 completed successfully. Download summary: %s", phase2_summary)
            run_summary = build_run_summary(
                keyword_filter=FILTER_KEYWORD,
                firm_name=FIRM_NAME,
                record_count=len(phase2_matches),
                outcomes=phase2_outcomes,
                issues=phase2_issues,
                elapsed_seconds=time.time() - run_started_at,
            )
            run_summary_path = save_run_summary(output_dir, run_summary)
            log.info("Run summary saved: %s", run_summary_path)
            if phase2_issues:
                log.warning(
                    "Phase 2 completed with %d logged issue(s). Review: %s",
                    len(phase2_issues),
                    phase2_issue_log,
                )
            else:
                log.info("Phase 2 completed with no logged download/naming issues.")
        elif matches:
            log.info("No phase-2 records left after applying resume filters.")
            run_summary = build_run_summary(
                keyword_filter=FILTER_KEYWORD,
                firm_name=FIRM_NAME,
                record_count=0,
                outcomes=[],
                issues=[],
                elapsed_seconds=time.time() - run_started_at,
            )
            save_run_summary(output_dir, run_summary)

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
