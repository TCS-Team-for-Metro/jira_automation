from __future__ import annotations

import sys
from csv import reader
from datetime import datetime
import json
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from uuid import uuid4

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

JIRA_ISSUES_URL = "https://jira.metro.digital/login.jsp"
CSV_DOWNLOAD_URL = "https://jira.metro.digital/sr/jira.issueviews:searchrequest-csv-current-fields/139211/SearchRequest-139211.csv"
MICROSOFT_EMAIL = "ashish.dake@metro-external.digital"

# power automate flow : Jira unlabeled 30 minute python flow
TEAMS_WEBHOOK_URL = "https://default6432230809a947a38c1cb82871d605.68.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/19/workflows/5ba2cd2ba42f4ba78c055e0365a07517/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=O79ugUWu8bMoOkKWHcER42xdxnHVzx9eJiWY--eOTXc"

def build_download_filename() -> str:
    print("Building CSV download filename.")
    today = datetime.now().strftime("%d-%b-%Y")
    return f"Wave 3 Daily Tickets {today}.csv"


def is_jira_url(url: str) -> bool:
    print(f"Checking whether URL is Jira: {url}")
    return "jira.metro.digital" in url


def is_microsoft_login_url(url: str) -> bool:
    print(f"Checking whether URL is a Microsoft login page: {url}")
    return "login.microsoftonline.com" in url or "login.live.com" in url


def is_effectively_empty_row(row: list[str]) -> bool:
    return all(not value.strip() for value in row)


def fill_microsoft_email(page) -> bool:
    print("Attempting to fill Microsoft email field.")
    selectors = [
        "input[type='email'][name='loginfmt']",
        "input[name='loginfmt']",
        "input[type='email']",
        "input#i0116",
        "input[type='text'][name='loginfmt']",
        "input[type='text']",
    ]

    for selector in selectors:
        locator = page.locator(selector)
        if locator.count() == 0:
            continue

        field = locator.first
        if not field.is_visible():
            continue

        field.click()
        field.press("Control+A")
        field.fill(MICROSOFT_EMAIL)
        print(f"Microsoft email field filled using selector: {selector}")
        return True

    print("Microsoft email field is not visible.")
    return False


def click_primary_button(page, *texts: str):
    print(f"Trying primary buttons: {texts}")
    for text in texts:
        button = page.get_by_role("button", name=text)
        if button.count() > 0 and button.first.is_visible():
            button.first.click()
            print(f"Clicked button with text: {text}")
            return True
    return False


def login_to_microsoft(page) -> None:
    print("Starting Jira login flow.")
    page.goto(JIRA_ISSUES_URL, wait_until="domcontentloaded", timeout=120000)

    deadline = datetime.now().timestamp() + 120
    while datetime.now().timestamp() < deadline:
        current_url = page.url.lower()
        if is_jira_url(current_url):
            print("Jira page detected. Login flow complete.")
            return

        if is_microsoft_login_url(current_url):
            if fill_microsoft_email(page):
                if not click_primary_button(page, "Next", "Continue"):
                    print("Falling back to default submit button.")
                    submit_buttons = page.locator("input[type='submit'], button[type='submit']")
                    if submit_buttons.count() > 0 and submit_buttons.first.is_visible():
                        submit_buttons.first.click()
                    else:
                        page.locator("input[type='submit']").first.click()
                page.wait_for_timeout(2000)
                continue

            print("Waiting for Microsoft login page to finish loading.")
            page.wait_for_load_state("domcontentloaded", timeout=15000)
            page.wait_for_timeout(1000)
            continue

        page.wait_for_timeout(1000)

    raise RuntimeError("Microsoft login flow did not reach the Jira dashboard.")


def download_csv(page) -> Path:
    print("Starting CSV download.")
    download_dir = Path.home() / "Downloads"
    download_dir.mkdir(parents=True, exist_ok=True)

    file_name = build_download_filename()
    temp_file = download_dir / f"{uuid4()}-{file_name}"

    with page.expect_download() as download_info:
        page.evaluate("url => window.location.href = url", CSV_DOWNLOAD_URL)

    download = download_info.value
    download.save_as(path=temp_file)

    final_path = download_dir / file_name
    temp_file.replace(final_path)
    print(f"CSV downloaded to: {final_path}")
    return final_path


def csv_has_data_rows(csv_path: Path) -> bool:
    print(f"Checking CSV for data rows: {csv_path}")
    with csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        for row_number, row in enumerate(reader(csv_file), start=1):
            if row_number == 1:
                continue
            if is_effectively_empty_row(row):
                continue
            return True
    return False


def extract_ticket_details(csv_path: Path) -> list[dict[str, str]]:
    print(f"Extracting ticket details from CSV: {csv_path}")
    with csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        rows = list(reader(csv_file))

    if not rows:
        return []

    headers = [header.strip() for header in rows[0]]
    ticket_details: list[dict[str, str]] = []
    for row in rows[1:]:
        if is_effectively_empty_row(row):
            continue

        row_values = list(row) + [""] * max(0, len(headers) - len(row))
        ticket_details.append(
            {header: row_values[index].strip() for index, header in enumerate(headers)}
        )

    return ticket_details


def send_ticket_details_to_teams_webhook(ticket_details: list[dict[str, str]]) -> None:
    print("Sending ticket details to Teams webhook.")
    payload = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "ticket_details": ticket_details,
    }
    request = Request(
        TEAMS_WEBHOOK_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=30):
            pass
    except HTTPError as exc:
        raise RuntimeError(
            f"Teams webhook returned HTTP {exc.code}: {exc.reason}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(f"Teams webhook request failed: {exc.reason}") from exc


def main() -> int:
    print("Launching browser automation.")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        csv_path: Path | None = None

        try:
            login_to_microsoft(page)
            csv_path = download_csv(page)
            has_data_rows = csv_has_data_rows(csv_path)
            if has_data_rows:
                print("CSV contains data rows.")
                ticket_details = extract_ticket_details(csv_path)
                send_ticket_details_to_teams_webhook(ticket_details)
            else:
                print("CSV does not contain data rows.")
            return 0
        except PlaywrightError as exc:
            print(f"Playwright error: {exc}", file=sys.stderr)
            return 1
        except Exception as exc:  # pragma: no cover
            print(f"Failed to download Jira CSV: {exc}", file=sys.stderr)
            return 1
        finally:
            if csv_path is not None and csv_path.exists():
                print(f"Removing downloaded CSV file: {csv_path}")
                # csv_path.unlink()
            print("Closing browser.")
            browser.close()


if __name__ == "__main__":
    raise SystemExit(main())
