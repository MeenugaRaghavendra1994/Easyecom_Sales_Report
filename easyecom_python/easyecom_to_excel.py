"""
EasyEcom MINI_SALES_REPORT -> Google Sheets automation

Flow:
1. POST /getApiToken
2. POST /reports/queue
3. GET /reports/download?reportId=...
4. Download ZIP (or ZIP URL returned as JSON)
5. Extract CSV/XLSX files and write combined data to Google Sheet "RAW"

Install:
    pip install -r requirements.txt

Configure credentials in environment variables or .env:
    EASYECOM_API_KEY
    EASYECOM_EMAIL
    EASYECOM_PASSWORD
    GOOGLE_SHEET_ID
    GOOGLE_SERVICE_ACCOUNT_FILE

Share the spreadsheet with the service-account email from the JSON file.

Run:
    python easyecom_to_excel.py
"""

import io
import json
import logging
import os
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import gspread
from dotenv import load_dotenv


# ------------------------- CONFIGURATION -------------------------

load_dotenv()

API_TOKEN_URL = "https://api.easyecom.io/getApiToken"
REPORT_QUEUE_URL = "https://api.easyecom.io/reports/queue"
REPORT_DOWNLOAD_URL = "https://api.easyecom.io/reports/download"

REPORT_TYPE = "MINI_SALES_REPORT"
INVOICE_TYPE = "ALL"
WAREHOUSE_IDS = "ix82194596416"
DATE_TYPE = "ORDER_DATE"
START_DATE = "2026-10-01"
END_DATE = datetime.now(ZoneInfo("Asia/Kolkata")).date().isoformat()

SHEET_NAME = "RAW"
GOOGLE_SHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID", "19HFavfdydCP3FnBkdSSCTulha4OUTY3X1Dl7yHrVWMA"
)
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv(
    "GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json"
)
PENDING_REPORT_ID_FILE = Path("easyecom_pending_report_id.txt")

MAX_DOWNLOAD_ATTEMPTS = 40
RETRY_SECONDS = 15
REQUEST_TIMEOUT = 90

# Optional: set a custom User-Agent if your successful Postman request uses one.
# Do not use this to bypass CAPTCHA/WAF. If WAF blocks Python, ask EasyEcom
# to allow the integration's server-to-server traffic.
USER_AGENT = os.getenv("EASYECOM_USER_AGENT", "EasyEcom-Report-Integration/1.0")


# ------------------------- LOGGING -------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("easyecom_execution.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("easyecom")


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(
            f"Missing {name}. Add it to your .env file or environment variables."
        )
    return value


def response_excerpt(response: requests.Response, limit: int = 1500) -> str:
    """Return a safe response excerpt for logs; never log credentials or JWT."""
    content_type = response.headers.get("Content-Type", "")
    if "json" in content_type.lower():
        try:
            body = response.json()
            # Avoid accidentally logging token values returned by API-1.
            if isinstance(body, dict) and isinstance(body.get("data"), dict):
                body["data"].pop("jwt_token", None)
                body["data"].pop("api_token", None)
            return json.dumps(body, ensure_ascii=False)[:limit]
        except ValueError:
            pass
    return response.text[:limit]


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "Accept": "application/json",
            "User-Agent": USER_AGENT,
        }
    )
    return session


def get_jwt_token(session: requests.Session, api_key: str,
                  email: str, password: str) -> str:
    log.info("STEP 1: Requesting EasyEcom JWT token")
    response = session.post(
        API_TOKEN_URL,
        headers={"x-api-key": api_key, "Content-Type": "application/json"},
        json={"email": email, "password": password},
        timeout=REQUEST_TIMEOUT,
    )
    log.info("API-1 HTTP status: %s", response.status_code)

    response.raise_for_status()
    data = response.json()
    token = (data.get("data") or {}).get("jwt_token")
    if not token:
        raise RuntimeError(
            "API-1 returned HTTP success, but data.jwt_token was not found. "
            "Check the API-1 response structure."
        )
    log.info("JWT received. Token length: %s (token value not logged)", len(token))
    return token


def generate_report(session: requests.Session, api_key: str,
                    jwt_token: str) -> dict:
    log.info("STEP 2: Queueing MINI_SALES_REPORT")
    payload = {
        "reportType": REPORT_TYPE,
        "params": {
            "invoiceType": INVOICE_TYPE,
            "warehouseIds": WAREHOUSE_IDS,
            "dateType": DATE_TYPE,
            "startDate": START_DATE,
            "endDate": END_DATE,
        },
    }
    headers = {
        "x-api-key": api_key,
        "Authorization": f"Bearer {jwt_token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    log.info("API-2 request URL: %s", REPORT_QUEUE_URL)
    log.info("API-2 request payload: %s", json.dumps(payload))
    log.info("API-2 Authorization: Bearer JWT returned by API-1")
    response = session.post(
        REPORT_QUEUE_URL,
        headers=headers,
        json=payload,
        timeout=REQUEST_TIMEOUT,
    )
    log.info("API-2 HTTP status: %s", response.status_code)
    log.info("API-2 response: %s", response_excerpt(response))

    if response.status_code == 405 and (
        "human verification" in response.text.lower()
        or "awswaf" in response.text.lower()
    ):
        raise RuntimeError(
            "API-2 was blocked by AWS WAF CAPTCHA (HTTP 405). "
            "This is not a normal JWT error. Since the request works in Postman, "
            "compare Postman's generated cURL with this Python request. "
            "If Python is also challenged, ask EasyEcom to permit server-to-server "
            "requests from this machine/network. Do not attempt to bypass CAPTCHA."
        )

    response.raise_for_status()
    return response.json()


def find_value_recursive(obj, accepted_keys):
    if isinstance(obj, dict):
        for key in accepted_keys:
            value = obj.get(key)
            if value not in (None, ""):
                return value
        for value in obj.values():
            found = find_value_recursive(value, accepted_keys)
            if found not in (None, ""):
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = find_value_recursive(item, accepted_keys)
            if found not in (None, ""):
                return found
    return None


def get_report_id(report_response: dict) -> str:
    report_id = find_value_recursive(
        report_response, {"reportId", "report_id", "reportID"}
    )
    if report_id is None:
        raise RuntimeError(
            "Could not find reportId/report_id in API-2 response. "
            "Inspect the API-2 response logged above and update get_report_id()."
        )
    log.info("Report ID received: %s", report_id)
    return str(report_id)


def get_or_queue_report_id(session: requests.Session, api_key: str,
                           jwt_token: str) -> str:
    if PENDING_REPORT_ID_FILE.is_file():
        report_id = PENDING_REPORT_ID_FILE.read_text(encoding="utf-8").strip()
        if report_id:
            log.info("Resuming existing report ID: %s", report_id)
            return report_id

    report_response = generate_report(session, api_key, jwt_token)
    report_id = get_report_id(report_response)
    PENDING_REPORT_ID_FILE.write_text(report_id, encoding="utf-8")
    return report_id


def find_download_url(obj):
    if isinstance(obj, dict):
        for key, value in obj.items():
            if (
                isinstance(value, str)
                and value.lower().startswith(("https://", "http://"))
                and any(word in key.lower() for word in ("url", "download", "file", "link"))
            ):
                return value
        for value in obj.values():
            found = find_download_url(value)
            if found:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = find_download_url(item)
            if found:
                return found
    return None


def is_zip_bytes(content: bytes) -> bool:
    return len(content) >= 4 and content[:2] == b"PK"


def download_zip(session: requests.Session, api_key: str,
                 jwt_token: str, report_id: str) -> bytes:
    log.info("STEP 3: Downloading report ID %s", report_id)
    url = f"{REPORT_DOWNLOAD_URL}?reportId={report_id}"
    headers = {
        "x-api-key": api_key,
        "Authorization": f"Bearer {jwt_token}",
        "Accept": "*/*",
    }

    for attempt in range(1, MAX_DOWNLOAD_ATTEMPTS + 1):
        log.info("API-3 download attempt %s/%s", attempt, MAX_DOWNLOAD_ATTEMPTS)
        response = session.get(
            url,
            headers=headers,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
        )
        log.info("API-3 HTTP status: %s", response.status_code)
        log.info("API-3 Content-Type: %s", response.headers.get("Content-Type", ""))
        response.raise_for_status()

        if is_zip_bytes(response.content):
            log.info("API-3 returned ZIP bytes (%s bytes)", len(response.content))
            return response.content

        if "text/html" in response.headers.get("Content-Type", "").lower():
            raise RuntimeError(
                "API-3 returned an HTML page instead of report data. Verify the "
                "download endpoint and authentication with EasyEcom."
            )

        # API-3 may instead return JSON containing a signed ZIP URL.
        try:
            body = response.json()
        except ValueError:
            body = None

        if isinstance(body, (dict, list)):
            download_url = find_download_url(body)
            data = body.get("data", {}) if isinstance(body, dict) else {}
            report_status = data.get("reportStatus", "unknown") if isinstance(data, dict) else "unknown"
            log.info(
                "API-3 report status: %s | download URL present: %s",
                report_status, bool(download_url),
            )
            if download_url:
                log.info("ZIP download URL found in API-3 response")
                # Signed object-storage URLs commonly do not require EasyEcom headers.
                file_response = session.get(
                    download_url,
                    timeout=REQUEST_TIMEOUT,
                    allow_redirects=True,
                )
                log.info("ZIP URL HTTP status: %s", file_response.status_code)
                file_response.raise_for_status()
                if not is_zip_bytes(file_response.content):
                    raise RuntimeError(
                        "The download URL responded successfully, but the content "
                        "does not appear to be a ZIP file."
                    )
                log.info("ZIP downloaded (%s bytes)", len(file_response.content))
                return file_response.content

        # Non-ZIP response may mean report is still processing.
        log.warning(
            "Report is not ready or response is not ZIP (HTTP %s, Content-Type: %s)",
            response.status_code, response.headers.get("Content-Type", ""),
        )
        if attempt < MAX_DOWNLOAD_ATTEMPTS:
            time.sleep(RETRY_SECONDS)

    raise TimeoutError(
        "Report did not become downloadable after "
        f"{MAX_DOWNLOAD_ATTEMPTS} attempts."
    )


def read_zip_tables(zip_bytes: bytes) -> pd.DataFrame:
    log.info("STEP 4: Extracting report ZIP in memory")
    frames = []

    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as archive:
            names = archive.namelist()
            log.info("ZIP contains %s entries: %s", len(names), names)

            for name in names:
                if name.endswith("/") or name.startswith("__MACOSX/"):
                    continue

                suffix = Path(name).suffix.lower()
                raw = archive.read(name)

                if suffix == ".csv":
                    log.info("Reading CSV: %s (%s bytes)", name, len(raw))
                    parsed = None
                    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
                        try:
                            parsed = pd.read_csv(
                                io.BytesIO(raw), dtype=str, encoding=encoding,
                                keep_default_na=False
                            )
                            break
                        except UnicodeDecodeError:
                            continue
                    if parsed is None:
                        raise RuntimeError(f"Could not decode CSV file: {name}")
                    frames.append(parsed)

                elif suffix in (".xlsx", ".xls"):
                    log.info("Reading Excel file inside ZIP: %s", name)
                    frames.append(
                        pd.read_excel(io.BytesIO(raw), dtype=str).fillna("")
                    )
                else:
                    log.info("Skipping unsupported ZIP entry: %s", name)

    except zipfile.BadZipFile as exc:
        raise RuntimeError("API-3 download is not a valid ZIP archive.") from exc

    if not frames:
        raise RuntimeError(
            "ZIP was extracted, but no CSV/XLSX files were found inside it."
        )

    # Combine multiple files with the same report columns.
    combined = pd.concat(frames, ignore_index=True, sort=False).fillna("")
    log.info("Combined report rows: %s | columns: %s",
             len(combined), len(combined.columns))
    return combined


def save_to_google_sheet(df: pd.DataFrame):
    credentials_path = Path(GOOGLE_SERVICE_ACCOUNT_FILE).expanduser()
    if not credentials_path.is_file():
        raise RuntimeError(
            "Google service-account JSON file not found: "
            f"{credentials_path.resolve()}. Set GOOGLE_SERVICE_ACCOUNT_FILE."
        )

    selected_column_indexes = [4, 6, 10, 11, 13, 20, 22, 26, 27, 28, 31, 32, 36, 37]
    if df.shape[1] <= max(selected_column_indexes):
        raise RuntimeError(
            "EasyEcom report has fewer than 38 columns; cannot apply the "
            "configured Google Sheets columns and filters."
        )

    clean_df = df.copy()
    for column_index in selected_column_indexes:
        clean_df.iloc[:, column_index] = (
            clean_df.iloc[:, column_index]
            .fillna("")
            .astype(str)
            .str.replace("`", "", regex=False)
        )

    channel = clean_df.iloc[:, 6].str.strip()
    status = clean_df.iloc[:, 10].str.strip()
    accepted_statuses = {"Shipped", "Assigned", "Confirmed", "Pending"}
    filtered_df = clean_df.loc[
        channel.eq("B2C") & status.isin(accepted_statuses)
    ].iloc[:, selected_column_indexes]
    filtered_df["Updated AT"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log.info("STEP 5: Updating Google Sheet tab %s", SHEET_NAME)
    log.info(
        "Filtered report rows: %s -> %s; selected columns: %s",
        len(df), len(filtered_df), len(selected_column_indexes),
    )
    client = gspread.service_account(filename=str(credentials_path))
    worksheet = client.open_by_key(GOOGLE_SHEET_ID).worksheet(SHEET_NAME)
    data = filtered_df.fillna("").astype(str).values.tolist()
    headers = [str(column) for column in filtered_df.columns]

    worksheet.clear()
    worksheet.resize(rows=max(len(data) + 1, 1), cols=max(len(headers), 1))
    worksheet.update(
        values=[headers], range_name="A1", value_input_option="RAW"
    )

    chunk_size = 100
    for offset in range(0, len(data), chunk_size):
        worksheet.update(
            values=data[offset:offset + chunk_size],
            range_name=f"A{offset + 2}",
            value_input_option="RAW",
        )

    log.info(
        "Google Sheet updated. Rows: %s | Columns: %s | Spreadsheet ID: %s",
        len(data), len(headers), GOOGLE_SHEET_ID,
    )


def main():
    started = time.time()
    log.info("=" * 65)
    log.info("EASYECOM AUTOMATION STARTED")
    log.info("Date range: %s to %s", START_DATE, END_DATE)
    log.info("Warehouse IDs: %s", WAREHOUSE_IDS)
    log.info("=" * 65)

    api_key = required_env("EASYECOM_API_KEY")
    email = required_env("EASYECOM_EMAIL")
    password = required_env("EASYECOM_PASSWORD")

    session = make_session()
    jwt_token = get_jwt_token(session, api_key, email, password)
    report_id = get_or_queue_report_id(session, api_key, jwt_token)
    zip_bytes = download_zip(session, api_key, jwt_token, report_id)
    report_df = read_zip_tables(zip_bytes)
    save_to_google_sheet(report_df)
    PENDING_REPORT_ID_FILE.unlink(missing_ok=True)

    log.info("=" * 65)
    log.info("AUTOMATION COMPLETED in %.1f seconds", time.time() - started)
    log.info("=" * 65)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log.exception("AUTOMATION FAILED")
        raise
