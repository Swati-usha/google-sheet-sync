import json
import logging
import os
import sys

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook, load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

GOOGLE_SHEET_NAME = "Birthday Memo"
GOOGLE_WORKSHEET_NAME = "Form Responses 1"

CREDENTIALS_FILE = "credentials.json"
WATERMARK_FILE = "watermark.json"
EXCEL_FILE = "Birthdays.xlsx"

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s | %(levelname)s | %(message)s",
)

logger = logging.getLogger(__name__)


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

def connect_to_google_sheet():
    """Connect to Google Sheets using service-account credentials."""

    logger.info("Connecting to Google Sheets...")

    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f"Credentials file not found: {CREDENTIALS_FILE}"
        )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]

    credentials = Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=scopes,
    )

    client = gspread.authorize(credentials)

    spreadsheet = client.open(GOOGLE_SHEET_NAME)

    worksheet = spreadsheet.worksheet(
        GOOGLE_WORKSHEET_NAME
    )

    logger.info(
        "Connected to '%s' → '%s'",
        GOOGLE_SHEET_NAME,
        GOOGLE_WORKSHEET_NAME,
    )

    return worksheet


# ============================================================
# WATERMARK
# ============================================================

def read_watermark():
    """Read the last successfully processed Google Sheet row."""

    if not os.path.exists(WATERMARK_FILE):
        logger.warning(
            "%s not found. Starting from row 1.",
            WATERMARK_FILE,
        )

        return 1

    try:

        with open(WATERMARK_FILE, "r") as file:
            data = json.load(file)

        last_row = int(data.get("last_row", 1))

        if last_row < 1:
            raise ValueError(
                "Watermark must be >= 1."
            )

        return last_row

    except (json.JSONDecodeError, ValueError, TypeError) as error:

        raise ValueError(
            f"Invalid watermark file: {error}"
        )


def write_watermark(last_row):
    """Persist watermark only after Excel update succeeds."""

    temporary_file = f"{WATERMARK_FILE}.tmp"

    with open(
        temporary_file,
        "w",
    ) as file:

        json.dump(
            {"last_row": last_row},
            file,
            indent=4,
        )

    # Atomic replacement
    os.replace(
        temporary_file,
        WATERMARK_FILE,
    )

    logger.info(
        "Watermark updated to row %s",
        last_row,
    )


# ============================================================
# EXTRACT
# ============================================================

def get_new_rows(worksheet, last_row):
    """
    Read the Google Sheet and return rows after the watermark.

    The watermark represents the physical Google Sheet row
    that was successfully written to Excel.
    """

    logger.info(
        "Last processed row: %s",
        last_row,
    )

    values = worksheet.get_all_values()

    actual_last_row = len(values)

    logger.info(
        "Actual Google Sheet rows: %s",
        actual_last_row,
    )

    # --------------------------------------------------------
    # Validate watermark
    # --------------------------------------------------------

    if last_row > actual_last_row:

        raise ValueError(
            f"Watermark ({last_row}) is ahead of "
            f"Google Sheet ({actual_last_row}). "
            "Manual intervention required."
        )

    # --------------------------------------------------------
    # No new records
    # --------------------------------------------------------

    if actual_last_row == last_row:

        logger.info("No new rows found.")

        return [], last_row

    # --------------------------------------------------------
    # Extract new rows
    # --------------------------------------------------------

    new_rows = values[last_row:]

    new_last_row = actual_last_row

    logger.info(
        "Fetched %s new rows.",
        len(new_rows),
    )

    logger.info(
        "New last row: %s",
        new_last_row,
    )

    return new_rows, new_last_row


# ============================================================
# EXCEL
# ============================================================

def create_excel(headers):
    """Create a new Excel workbook."""

    logger.info(
        "Creating new Excel file: %s",
        EXCEL_FILE,
    )

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Responses"

    worksheet.append(headers)

    workbook.save(EXCEL_FILE)

    return workbook, worksheet


def append_to_excel(headers, new_rows):
    """Append new Google Sheet rows to Excel."""

    if not new_rows:
        return

    # --------------------------------------------------------
    # Create Excel if it doesn't exist
    # --------------------------------------------------------

    if not os.path.exists(EXCEL_FILE):

        workbook, worksheet = create_excel(headers)

    else:

        logger.info(
            "Opening existing Excel file: %s",
            EXCEL_FILE,
        )

        workbook = load_workbook(
            EXCEL_FILE
        )

        if "Responses" in workbook.sheetnames:

            worksheet = workbook["Responses"]

        else:

            worksheet = workbook.create_sheet(
                "Responses"
            )

            worksheet.append(headers)

    # --------------------------------------------------------
    # Append records
    # --------------------------------------------------------

    for row in new_rows:

        worksheet.append(row)

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    workbook.save(EXCEL_FILE)

    workbook.close()

    logger.info(
        "Appended %s rows to Excel.",
        len(new_rows),
    )


# ============================================================
# MAIN
# ============================================================

def main():

    logger.info(
        "========== Birthday Memo Sync Started =========="
    )

    try:

        # ----------------------------------------------------
        # 1. Read watermark
        # ----------------------------------------------------

        last_row = read_watermark()

        # ----------------------------------------------------
        # 2. Connect to Google Sheet
        # ----------------------------------------------------

        worksheet = connect_to_google_sheet()

        # ----------------------------------------------------
        # 3. Read headers
        # ----------------------------------------------------

        all_values = worksheet.get_all_values()

        if not all_values:

            raise ValueError(
                "Google Sheet is empty."
            )

        headers = all_values[0]

        logger.info(
            "Detected %s columns.",
            len(headers),
        )

        # ----------------------------------------------------
        # 4. Extract new rows
        # ----------------------------------------------------

        new_rows, new_last_row = get_new_rows(
            worksheet,
            last_row,
        )

        # ----------------------------------------------------
        # 5. Nothing to process
        # ----------------------------------------------------

        if not new_rows:

            logger.info(
                "========== Sync Completed: Nothing to Update =========="
            )

            return

        # ----------------------------------------------------
        # 6. Load into Excel
        # ----------------------------------------------------

        append_to_excel(
            headers,
            new_rows,
        )

        # ----------------------------------------------------
        # 7. Update watermark ONLY after successful load
        # ----------------------------------------------------

        write_watermark(
            new_last_row
        )

        logger.info(
            "========== Sync Completed Successfully =========="
        )

    except Exception as error:

        logger.exception(
            "SYNC FAILED: %s",
            error,
        )

        # Non-zero exit code is important for GitHub Actions.
        sys.exit(1)


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()