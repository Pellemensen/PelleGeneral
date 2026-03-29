"""
50/30/20 Budget Tracker — Google Sheets builder
Run:  python budget_tracker.py
Requires:
  pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
  A credentials.json (OAuth 2.0) or service_account.json in the same directory.
"""

from __future__ import annotations
import json, os, sys, time
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_TITLE = "50/30/20 Budget Tracker"

# ── Colour helpers ────────────────────────────────────────────────────────────

def c(h: str) -> dict:
    """Hex → Sheets RGB dict (0–1 floats)."""
    h = h.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255, "green": g / 255, "blue": b / 255}

SAGE        = c("5C8C6E")
LIGHT_SAGE  = c("7AAD8A")
DUSTY_ROSE  = c("D98C8C")
WARM_SAND   = c("C4A882")
SOFT_MINT   = c("E8F0EB")
ALT_ROW     = c("F5FAF7")
OFF_WHITE   = c("F9F5F0")
DARK_TEXT   = c("2C3E2D")
WHITE       = {"red": 1, "green": 1, "blue": 1}
RED_BG      = c("FADBD8")
RED_TEXT    = c("C0392B")
GREEN_BG    = c("D5F5E3")
GREEN_TEXT  = c("1E8449")
YELLOW_BG   = c("FEF9E7")
GREY_BORDER = c("BDC3C7")

# ── Grid helpers ──────────────────────────────────────────────────────────────

def grid(sheet_id: int, r1: int, c1: int, r2: int, c2: int) -> dict:
    return {
        "sheetId": sheet_id,
        "startRowIndex": r1, "endRowIndex": r2,
        "startColumnIndex": c1, "endColumnIndex": c2,
    }


def cell_fmt(sheet_id, row, col, **props) -> dict:
    return repeat_fmt(sheet_id, row, col, row + 1, col + 1, **props)


def repeat_fmt(sheet_id, r1, c1, r2, c2, **props) -> dict:
    fields = ",".join(f"userEnteredFormat.{k}" for k in props)
    fmt = {}
    for k, v in props.items():
        fmt[k] = v
    return {
        "repeatCell": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": fmt},
            "fields": f"userEnteredFormat({','.join(props.keys())})",
        }
    }


def bg(sheet_id, r1, c1, r2, c2, colour: dict) -> dict:
    return {
        "repeatCell": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": {"backgroundColor": colour}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    }


def font(sheet_id, r1, c1, r2, c2, size=9, bold=False, colour=None, italic=False) -> dict:
    tf = {"fontSize": size, "bold": bold, "italic": italic}
    if colour:
        tf["foregroundColor"] = colour
    return {
        "repeatCell": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": {"textFormat": tf}},
            "fields": "userEnteredFormat.textFormat",
        }
    }


def align(sheet_id, r1, c1, r2, c2, h="LEFT", v="MIDDLE") -> dict:
    return {
        "repeatCell": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": h,
                "verticalAlignment": v,
            }},
            "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)",
        }
    }


def number_fmt(sheet_id, r1, c1, r2, c2, pattern: str) -> dict:
    return {
        "repeatCell": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": pattern}
            }},
            "fields": "userEnteredFormat.numberFormat",
        }
    }


def col_width(sheet_id, col_idx: int, px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col_idx,
                "endIndex": col_idx + 1,
            },
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def row_height(sheet_id, row_idx: int, px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": row_idx,
                "endIndex": row_idx + 1,
            },
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def merge(sheet_id, r1, c1, r2, c2) -> dict:
    return {"mergeCells": {"range": grid(sheet_id, r1, c1, r2, c2), "mergeType": "MERGE_ALL"}}


def freeze(sheet_id, rows=1, cols=0) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols},
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def borders_outer(sheet_id, r1, c1, r2, c2, colour=None, width=2) -> dict:
    colour = colour or SAGE
    side = {"style": "SOLID_MEDIUM" if width >= 2 else "SOLID", "color": colour}
    return {
        "updateBorders": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "top": side, "bottom": side, "left": side, "right": side,
        }
    }


def borders_inner(sheet_id, r1, c1, r2, c2) -> dict:
    side = {"style": "SOLID", "color": GREY_BORDER}
    return {
        "updateBorders": {
            "range": grid(sheet_id, r1, c1, r2, c2),
            "innerHorizontal": side, "innerVertical": side,
        }
    }


# ── Authentication ────────────────────────────────────────────────────────────

def get_service():
    creds = None
    base = os.path.dirname(os.path.abspath(__file__))

    # Try service account first
    sa_file = os.path.join(base, "service_account.json")
    if os.path.exists(sa_file):
        creds = service_account.Credentials.from_service_account_file(sa_file, scopes=SCOPES)
        return build("sheets", "v4", credentials=creds)

    # OAuth2 fallback
    token_path = os.path.join(base, "token.pickle")
    if os.path.exists(token_path):
        with open(token_path, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(base, "credentials.json"), SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)
    return build("sheets", "v4", credentials=creds)


# ── Sheet IDs ─────────────────────────────────────────────────────────────────
DASH = 0   # Dashboard
LOG  = 1   # Transactions Log
SET  = 2   # Settings
REF  = 3   # Reference (hidden)


# ── Category / bucket data ────────────────────────────────────────────────────

NEEDS_CATS = [
    "Internet", "Electricity", "Water", "Mobile", "Insurance",
    "Gas", "City Garbage", "Rent", "Groceries", "Transportation",
    "Household", "Pet",
]
WANTS_CATS = [
    "Apparel", "Dining Out", "Entertainment", "Travel", "Social Life",
    "Beauty", "Gift", "Netflix", "Amazon Prime", "Spotify",
    "Subscription Box", "Gym Membership",
]
SAVINGS_CATS = [
    "Emergency Fund", "Retirement", "Investment", "Credit Card",
    "Student Loan", "Car Loan", "Mortgage Overpayment", "Savings Goal",
]

INCOME_CATEGORIES = [
    "Paycheck", "Business", "Side Hustle", "Dividends",
    "Interest Income", "Commission",
    "Income 7", "Income 8", "Income 9", "Income 10", "Income 11", "Income 12",
]

ALL_EXPENSE_CATS = NEEDS_CATS + WANTS_CATS + SAVINGS_CATS

MONTH_OPTIONS = [
    "Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "May 2026", "Jun 2026",
    "Jul 2026", "Aug 2026", "Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026",
]


# ── Sample transactions ───────────────────────────────────────────────────────

SAMPLE_TRANSACTIONS = [
    # Jan income
    ("2026-01-01", "Paycheck",              "Income",  "Needs", "Paycheck",      1200),
    ("2026-01-01", "Business Revenue",      "Income",  "Needs", "Business",      2350),
    ("2026-01-15", "Side Hustle",           "Income",  "Needs", "Side Hustle",    800),
    ("2026-01-20", "Dividend",              "Income",  "Needs", "Dividends",      175),
    # Jan needs
    ("2026-01-05", "Rent payment",          "Expense", "Needs", "Rent",           725),
    ("2026-01-06", "Weekly groceries",      "Expense", "Needs", "Groceries",      204),
    ("2026-01-08", "Internet bill",         "Expense", "Needs", "Internet",        40),
    ("2026-01-09", "Electricity",           "Expense", "Needs", "Electricity",    200),
    ("2026-01-10", "Mobile plan",           "Expense", "Needs", "Mobile",          35),
    ("2026-01-12", "Car insurance",         "Expense", "Needs", "Insurance",      200),
    ("2026-01-14", "Gas",                   "Expense", "Needs", "Gas",            100),
    ("2026-01-18", "Uber to work",          "Expense", "Needs", "Transportation", 140),
    ("2026-01-22", "Household supplies",    "Expense", "Needs", "Household",      225),
    ("2026-01-25", "Pet food & vet",        "Expense", "Needs", "Pet",            150),
    # Jan wants
    ("2026-01-03", "New jacket",            "Expense", "Wants", "Apparel",         95),
    ("2026-01-07", "Restaurant dinner",     "Expense", "Wants", "Dining Out",     100),
    ("2026-01-11", "Cinema tickets",        "Expense", "Wants", "Entertainment",  100),
    ("2026-01-15", "Flight booking",        "Expense", "Wants", "Travel",         400),
    ("2026-01-17", "Bar with friends",      "Expense", "Wants", "Social Life",     90),
    ("2026-01-19", "Haircut",               "Expense", "Wants", "Beauty",         135),
    ("2026-01-23", "Birthday gift",         "Expense", "Wants", "Gift",            45),
    ("2026-01-25", "Netflix",               "Expense", "Wants", "Netflix",         35),
    ("2026-01-26", "Amazon Prime",          "Expense", "Wants", "Amazon Prime",    60),
    ("2026-01-28", "Spotify",               "Expense", "Wants", "Spotify",         15),
    ("2026-01-29", "Mystery Box",           "Expense", "Wants", "Subscription Box",120),
    ("2026-01-30", "Gym",                   "Expense", "Wants", "Gym Membership",  150),
    # Jan savings
    ("2026-01-02", "Emergency fund top-up", "Expense", "Savings & Debt", "Emergency Fund", 200),
    ("2026-01-02", "Retirement contribution","Expense","Savings & Debt", "Retirement",      300),
    ("2026-01-03", "Credit card payment",   "Expense", "Savings & Debt", "Credit Card",    360),
    # Feb
    ("2026-02-01", "Paycheck",              "Income",  "Needs", "Paycheck",      1200),
    ("2026-02-01", "Business Revenue",      "Income",  "Needs", "Business",      2100),
    ("2026-02-05", "Rent payment",          "Expense", "Needs", "Rent",           725),
    ("2026-02-07", "Weekly groceries",      "Expense", "Needs", "Groceries",      185),
    ("2026-02-10", "Restaurant dinner",     "Expense", "Wants", "Dining Out",      80),
    ("2026-02-14", "Valentine dinner",      "Expense", "Wants", "Dining Out",     120),
    ("2026-02-02", "Retirement contribution","Expense","Savings & Debt","Retirement",300),
    # Mar (sparse)
    ("2026-03-01", "Paycheck",              "Income",  "Needs", "Paycheck",      1200),
    ("2026-03-05", "Rent payment",          "Expense", "Needs", "Rent",           725),
    ("2026-03-08", "Cinema tickets",        "Expense", "Wants", "Entertainment",   65),
]


# ── Create spreadsheet skeleton ───────────────────────────────────────────────

def create_spreadsheet(service) -> str:
    body = {
        "properties": {"title": SPREADSHEET_TITLE},
        "sheets": [
            {"properties": {"sheetId": DASH, "title": "Dashboard",         "index": 0}},
            {"properties": {"sheetId": LOG,  "title": "Transactions Log",  "index": 1}},
            {"properties": {"sheetId": SET,  "title": "Settings",          "index": 2}},
            {"properties": {"sheetId": REF,  "title": "Reference",         "index": 3,
                             "hidden": True}},
        ],
    }
    res = service.spreadsheets().create(body=body, fields="spreadsheetId").execute()
    return res["spreadsheetId"]


def batch(service, spreadsheet_id: str, requests: list):
    if not requests:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()


def values_update(service, sid: str, range_: str, rows: list):
    service.spreadsheets().values().update(
        spreadsheetId=sid,
        range=range_,
        valueInputOption="USER_ENTERED",
        body={"values": rows},
    ).execute()


# ── Reference tab ─────────────────────────────────────────────────────────────

def setup_reference(service, sid):
    months = [["January"], ["February"], ["March"], ["April"],
              ["May"], ["June"], ["July"], ["August"],
              ["September"], ["October"], ["November"], ["December"]]
    buckets = [["Needs"], ["Wants"], ["Savings & Debt"]]
    statuses = [["On Track"], ["Over Budget"]]
    values_update(service, sid, "Reference!A1:A12", months)
    values_update(service, sid, "Reference!B1:B3",  buckets)
    values_update(service, sid, "Reference!C1:C2",  statuses)


# ── Settings tab ──────────────────────────────────────────────────────────────

def setup_settings(service, sid):
    # Ratio header + values
    header_rows = [
        ["BUDGET RATIOS", ""],
        ["Needs %",       50],
        ["Wants %",       30],
        ["Savings & Debt %", 20],
        ["Ratio check",   '=IF(B2+B3+B4=100,"✓ Ratios OK","⚠ Must sum to 100%")'],
    ]
    values_update(service, sid, "Settings!A1:B5", header_rows)

    # Category → bucket mapping header
    values_update(service, sid, "Settings!A7:B7", [["Category Name", "Bucket"]])

    mapping_rows = (
        [[cat, "Needs"]          for cat in NEEDS_CATS]
        + [[cat, "Wants"]        for cat in WANTS_CATS]
        + [[cat, "Savings & Debt"] for cat in SAVINGS_CATS]
    )
    values_update(service, sid, f"Settings!A8:B{8 + len(mapping_rows) - 1}", mapping_rows)

    reqs = [
        # Title
        bg(SET, 0, 0, 1, 2, SAGE),
        font(SET, 0, 0, 1, 2, size=11, bold=True, colour=WHITE),
        # Section header row 7
        bg(SET, 6, 0, 7, 2, SAGE),
        font(SET, 6, 0, 7, 2, size=9, bold=True, colour=WHITE),
        # Alt rows for mapping
        freeze(SET, rows=1),
        col_width(SET, 0, 180),
        col_width(SET, 1, 130),
    ]
    # Alt row colours
    for i, _ in enumerate(mapping_rows):
        colour = WHITE if i % 2 == 0 else ALT_ROW
        reqs.append(bg(SET, 8 + i, 0, 9 + i, 2, colour))

    # Ratio validation green/red CF
    reqs += [
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid(SET, 4, 1, 5, 2)],
                    "booleanRule": {
                        "condition": {"type": "TEXT_CONTAINS", "values": [{"userEnteredValue": "✓"}]},
                        "format": {"backgroundColor": GREEN_BG,
                                   "textFormat": {"foregroundColor": GREEN_TEXT}},
                    },
                },
                "index": 0,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid(SET, 4, 1, 5, 2)],
                    "booleanRule": {
                        "condition": {"type": "TEXT_CONTAINS", "values": [{"userEnteredValue": "⚠"}]},
                        "format": {"backgroundColor": RED_BG,
                                   "textFormat": {"foregroundColor": RED_TEXT}},
                    },
                },
                "index": 1,
            }
        },
        # Number validation for ratios
        {
            "setDataValidation": {
                "range": grid(SET, 1, 1, 4, 2),
                "rule": {
                    "condition": {
                        "type": "NUMBER_BETWEEN",
                        "values": [{"userEnteredValue": "0"}, {"userEnteredValue": "100"}],
                    },
                    "showCustomUi": True,
                    "strict": False,
                },
            }
        },
    ]
    batch(service, sid, reqs)


# ── Transactions Log tab ──────────────────────────────────────────────────────

def setup_transactions_log(service, sid):
    # Header row
    headers = [["Date", "Description", "Type", "Bucket", "Category", "Amount", "Month", "Year", "Notes"]]
    values_update(service, sid, "Transactions Log!A1:I1", headers)

    # Sample data rows
    data_rows = []
    for tx in SAMPLE_TRANSACTIONS:
        date_str, desc, tx_type, bucket, cat, amount = tx
        # Month/Year helper formulas reference column A
        row_num = len(data_rows) + 2  # 1-indexed, row 1 is header
        month_f = f'=IF(A{row_num}="","",MONTH(A{row_num}))'
        year_f  = f'=IF(A{row_num}="","",YEAR(A{row_num}))'
        # Bucket auto-fill via VLOOKUP (but sample data has explicit bucket)
        bucket_f = f'=IFERROR(VLOOKUP(E{row_num},Settings!$A$8:$B$40,2,FALSE),"")'
        data_rows.append([date_str, desc, tx_type, bucket_f, cat, amount, month_f, year_f, ""])

    values_update(service, sid, f"Transactions Log!A2:I{len(data_rows)+1}", data_rows)

    reqs = [
        # Header styling
        bg(LOG, 0, 0, 1, 9, SAGE),
        font(LOG, 0, 0, 1, 9, size=9, bold=True, colour=WHITE),
        align(LOG, 0, 0, 1, 9, h="CENTER"),
        freeze(LOG, rows=1),
        # Column widths
        col_width(LOG, 0, 100),  # Date
        col_width(LOG, 1, 200),  # Description
        col_width(LOG, 2, 80),   # Type
        col_width(LOG, 3, 110),  # Bucket
        col_width(LOG, 4, 130),  # Category
        col_width(LOG, 5, 90),   # Amount
        col_width(LOG, 6, 60),   # Month
        col_width(LOG, 7, 60),   # Year
        col_width(LOG, 8, 200),  # Notes
        # Date format col A
        number_fmt(LOG, 1, 0, 200, 1, "DD MMM YYYY"),
        # Currency format col F
        number_fmt(LOG, 1, 5, 200, 6, '$#,##0.00'),
        # Alt rows
    ]
    # Alt rows for data
    for i in range(len(data_rows)):
        colour = WHITE if i % 2 == 0 else ALT_ROW
        reqs.append(bg(LOG, i + 1, 0, i + 2, 9, colour))

    # Data validation — Type dropdown (col C)
    reqs.append({
        "setDataValidation": {
            "range": grid(LOG, 1, 2, 200, 3),
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": "Income"}, {"userEnteredValue": "Expense"}],
                },
                "showCustomUi": True,
                "strict": True,
            },
        }
    })

    # Data validation — Category dropdown (col E) from Settings
    reqs.append({
        "setDataValidation": {
            "range": grid(LOG, 1, 4, 200, 5),
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": "=Settings!$A$8:$A$40"}],
                },
                "showCustomUi": True,
                "strict": True,
            },
        }
    })

    # Date validation col A
    reqs.append({
        "setDataValidation": {
            "range": grid(LOG, 1, 0, 200, 1),
            "rule": {
                "condition": {"type": "DATE_IS_VALID"},
                "showCustomUi": True,
                "strict": False,
            },
        }
    })

    # Colour scale on Amount col (F2:F200) for Expense rows
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(LOG, 1, 5, 200, 6)],
                "gradientRule": {
                    "minpoint": {"color": GREEN_BG,  "type": "MIN"},
                    "maxpoint": {"color": SAGE,       "type": "MAX"},
                },
            },
            "index": 0,
        }
    })

    batch(service, sid, reqs)


# ── Dashboard values and formulas ─────────────────────────────────────────────

# Shorthand for SUMIFS by bucket + month/year from B5
def sumifs_bucket(bucket: str) -> str:
    return (
        f"=IFERROR(SUMIFS('Transactions Log'!F:F,"
        f"'Transactions Log'!C:C,\"Expense\","
        f"'Transactions Log'!D:D,\"{bucket}\","
        f"'Transactions Log'!G:G,MONTH(DATEVALUE(\"1 \"&Dashboard!$B$5)),"
        f"'Transactions Log'!H:H,YEAR(DATEVALUE(\"1 \"&Dashboard!$B$5))),0)"
    )


def sumifs_cat(col_letter: str) -> str:
    """SUMIFS for a category cell in col_letter, filtered by selected month."""
    return (
        f"=IFERROR(SUMIFS('Transactions Log'!F:F,"
        f"'Transactions Log'!E:E,{col_letter},"
        f"'Transactions Log'!G:G,MONTH(DATEVALUE(\"1 \"&Dashboard!$B$5)),"
        f"'Transactions Log'!H:H,YEAR(DATEVALUE(\"1 \"&Dashboard!$B$5))),0)"
    )


def setup_dashboard_values(service, sid):
    """Write all static labels and formulas to Dashboard."""

    # ── Row 1: Title
    values_update(service, sid, "Dashboard!A1", [["50/30/20 BUDGET DASHBOARD"]])

    # ── Row 3-5: Controls
    values_update(service, sid, "Dashboard!A3:D3", [["BUDGET PERIOD", "01 Jan 2026", "to", "31 Jan 2026"]])
    values_update(service, sid, "Dashboard!A4:B4", [["SET CURRENCY", "$"]])
    values_update(service, sid, "Dashboard!A5:B5", [["SELECTED MONTH", "Jan 2026"]])

    # ── Row 7: KPI row labels + formulas
    # Columns: B=label, C=value | E=label, F=value | H=label, I=value
    values_update(service, sid, "Dashboard!B7:I7", [[
        "NEEDS",   sumifs_bucket("Needs"),   "",
        "WANTS",   sumifs_bucket("Wants"),   "",
        "SAVINGS & DEBT", sumifs_bucket("Savings & Debt"),
    ]])

    # ── Row 9: Section headers
    values_update(service, sid, "Dashboard!A9",  [["INCOME SUMMARY"]])
    values_update(service, sid, "Dashboard!F9",  [["NEEDS SUMMARY"]])
    values_update(service, sid, "Dashboard!M9",  [["WANTS SUMMARY"]])
    values_update(service, sid, "Dashboard!T9",  [["AMOUNT LEFT TO SPEND"]])

    # ── Row 10: Column headers — Income table (A–D)
    values_update(service, sid, "Dashboard!A10:D10",
                  [["Category", "Expected ($)", "Actual ($)", "Variance"]])

    # ── Row 10: Column headers — Needs table (F–K)
    values_update(service, sid, "Dashboard!F10:K10",
                  [["Category", "Due Date", "Expected ($)", "Actual ($)", "Progress %", "Progress Bar"]])

    # ── Row 10: Column headers — Wants table (M–R)
    values_update(service, sid, "Dashboard!M10:R10",
                  [["Category", "Due Date", "Expected ($)", "Actual ($)", "Progress %", "Progress Bar"]])

    # ── Rows 11–22: Income rows (12 categories, 0-indexed row 10 = header, 11–22 = data)
    income_rows = []
    for cat in INCOME_CATEGORIES:
        # Expected = user fills manually; Actual = SUMIFS income entries
        actual_f = (
            f"=IFERROR(SUMIFS('Transactions Log'!F:F,"
            f"'Transactions Log'!E:E,A{len(income_rows)+11},"
            f"'Transactions Log'!C:C,\"Income\","
            f"'Transactions Log'!G:G,MONTH(DATEVALUE(\"1 \"&$B$5)),"
            f"'Transactions Log'!H:H,YEAR(DATEVALUE(\"1 \"&$B$5))),0)"
        )
        income_rows.append([cat, 0, actual_f, f"=C{len(income_rows)+11}-B{len(income_rows)+11}"])

    values_update(service, sid, "Dashboard!A11:D22", income_rows)

    # ── Row 23: blank spacer already empty

    # ── Row 24: TOTAL INCOME
    values_update(service, sid, "Dashboard!A24:D24",
                  [["TOTAL INCOME",
                    "=SUM(B11:B22)",
                    "=SUM(C11:C22)",
                    "=SUM(D11:D22)"]])

    # ── Row 25: blank

    # ── Row 26: ROLLOVER (manual)
    values_update(service, sid, "Dashboard!A26:B26", [["ROLLOVER", 0]])

    # ── Row 27: TOTAL BUDGET
    values_update(service, sid, "Dashboard!A27:C27",
                  [["TOTAL BUDGET", "=C24+B26", ""]])

    # ── Needs table rows 11–22 (columns F–K, 0-indexed cols 5–10)
    # F=category, G=due date, H=expected, I=actual(sumifs), J=%, K=sparkline
    needs_rows = []
    for i, cat in enumerate(NEEDS_CATS):
        r = i + 11  # sheet row (1-indexed)
        actual_f = (
            f"=IFERROR(SUMIFS('Transactions Log'!F:F,"
            f"'Transactions Log'!E:E,F{r},"
            f"'Transactions Log'!G:G,MONTH(DATEVALUE(\"1 \"&$B$5)),"
            f"'Transactions Log'!H:H,YEAR(DATEVALUE(\"1 \"&$B$5))),0)"
        )
        pct_f  = f"=IFERROR(I{r}/H{r},0)"
        spark_f = (
            f'=IF(H{r}=0,"",SPARKLINE(I{r}/H{r},'
            f'{{"charttype","bar";"max",1;"color1","#5C8C6E";"color2","#F2DDD8"}}))'
        )
        needs_rows.append([cat, "", 0, actual_f, pct_f, spark_f])

    values_update(service, sid, f"Dashboard!F11:K{10+len(needs_rows)}", needs_rows)

    # Needs total row
    nr = 10 + len(needs_rows) + 1  # row after last needs row
    values_update(service, sid, f"Dashboard!F{nr}:K{nr}",
                  [["TOTAL NEEDS", "", f"=SUM(H11:H{nr-1})",
                    f"=SUM(I11:I{nr-1})", f"=IFERROR(I{nr}/H{nr},0)", ""]])

    # ── Wants table rows 11–22 (columns M–R, 0-indexed 12–17)
    wants_rows = []
    for i, cat in enumerate(WANTS_CATS):
        r = i + 11
        actual_f = (
            f"=IFERROR(SUMIFS('Transactions Log'!F:F,"
            f"'Transactions Log'!E:E,M{r},"
            f"'Transactions Log'!G:G,MONTH(DATEVALUE(\"1 \"&$B$5)),"
            f"'Transactions Log'!H:H,YEAR(DATEVALUE(\"1 \"&$B$5))),0)"
        )
        pct_f  = f"=IFERROR(P{r}/O{r},0)"
        spark_f = (
            f'=IF(O{r}=0,"",SPARKLINE(P{r}/O{r},'
            f'{{"charttype","bar";"max",1;"color1","#D98C8C";"color2","#F2DDD8"}}))'
        )
        wants_rows.append([cat, "", 0, actual_f, pct_f, spark_f])

    values_update(service, sid, f"Dashboard!M11:R{10+len(wants_rows)}", wants_rows)

    wr = 10 + len(wants_rows) + 1
    values_update(service, sid, f"Dashboard!M{wr}:R{wr}",
                  [["TOTAL WANTS", "", f"=SUM(O11:O{wr-1})",
                    f"=SUM(P11:P{wr-1})", f"=IFERROR(P{wr}/O{wr},0)", ""]])

    # ── Amount Left panel (cols T–W, 0-indexed 19–22), rows 10–15
    total_spent = "=SUM(I11:I23)+SUM(P11:P23)"  # needs + wants actuals
    amount_left = f"=C27-({total_spent})"
    pct_left    = f"=IFERROR(({amount_left})/C27,0)"
    status_msg  = (
        f'=IF(({pct_left})>0.1,'
        f'"✓ Doing great! You are right on track.",'
        f'"⚠ Watch your spending!")'
    )

    values_update(service, sid, "Dashboard!T10:W10", [["AMOUNT LEFT TO SPEND", "", "", ""]])
    values_update(service, sid, "Dashboard!T11",     [[amount_left]])
    values_update(service, sid, "Dashboard!T13",     [["PERCENTAGE LEFT"]])
    values_update(service, sid, "Dashboard!T14",     [[pct_left]])
    values_update(service, sid, "Dashboard!T15",     [[status_msg]])

    # ── Bucket targets (hidden helper — row 2, cols T–W)
    values_update(service, sid, "Dashboard!T2:V2", [[
        "=Settings!B2/100*C27",   # Needs target
        "=Settings!B3/100*C27",   # Wants target
        "=Settings!B4/100*C27",   # S&D target
    ]])


# ── Dashboard formatting ──────────────────────────────────────────────────────

def setup_dashboard_format(service, sid):
    reqs = [
        # Background fill
        bg(DASH, 0, 0, 100, 30, OFF_WHITE),

        # Title row merge + style
        merge(DASH, 0, 0, 1, 10),
        bg(DASH, 0, 0, 1, 10, SAGE),
        font(DASH, 0, 0, 1, 10, size=22, bold=True, colour=WHITE),
        align(DASH, 0, 0, 1, 10, h="CENTER", v="MIDDLE"),
        row_height(DASH, 0, 50),

        # Controls area background (rows 2–4)
        bg(DASH, 2, 0, 5, 10, SOFT_MINT),
        font(DASH, 2, 0, 5, 1, size=9, bold=True, colour=DARK_TEXT),
        align(DASH, 2, 0, 5, 10, v="MIDDLE"),

        # KPI row (row 6, 0-indexed)
        row_height(DASH, 6, 60),
        bg(DASH, 6, 1, 7, 10, SOFT_MINT),
        # KPI labels
        font(DASH, 6, 1, 7, 2, size=9,  bold=True, colour=DARK_TEXT),
        font(DASH, 6, 4, 7, 5, size=9,  bold=True, colour=DARK_TEXT),
        font(DASH, 6, 7, 7, 8, size=9,  bold=True, colour=DARK_TEXT),
        # KPI values
        font(DASH, 6, 2, 7, 3, size=18, bold=True, colour=DARK_TEXT),
        font(DASH, 6, 5, 7, 6, size=18, bold=True, colour=DARK_TEXT),
        font(DASH, 6, 8, 7, 9, size=18, bold=True, colour=DARK_TEXT),
        align(DASH, 6, 1, 7, 10, h="CENTER", v="MIDDLE"),
        number_fmt(DASH, 6, 2, 7, 3, '$#,##0.00'),
        number_fmt(DASH, 6, 5, 7, 6, '$#,##0.00'),
        number_fmt(DASH, 6, 8, 7, 9, '$#,##0.00'),

        # Section header rows (row 8, 0-indexed)
        bg(DASH, 8, 0, 9, 4,   SAGE),
        bg(DASH, 8, 5, 9, 11,  SAGE),
        bg(DASH, 8, 12, 9, 18, SAGE),
        bg(DASH, 8, 19, 9, 23, SAGE),
        font(DASH, 8, 0, 9, 23, size=10, bold=True, colour=WHITE),
        align(DASH, 8, 0, 9, 23, h="LEFT", v="MIDDLE"),

        # Column header rows (row 9, 0-indexed)
        bg(DASH, 9, 0, 10, 4,   LIGHT_SAGE),
        bg(DASH, 9, 5, 10, 11,  LIGHT_SAGE),
        bg(DASH, 9, 12, 10, 18, LIGHT_SAGE),
        bg(DASH, 9, 19, 10, 23, LIGHT_SAGE),
        font(DASH, 9, 0, 10, 23, size=9, bold=True, colour=WHITE),
        align(DASH, 9, 0, 10, 23, h="CENTER", v="MIDDLE"),

        # Column widths
        col_width(DASH, 0,  160),  # A
        col_width(DASH, 1,  90),   # B
        col_width(DASH, 2,  90),   # C
        col_width(DASH, 3,  80),   # D
        col_width(DASH, 4,  20),   # E spacer
        col_width(DASH, 5,  160),  # F
        col_width(DASH, 6,  70),   # G
        col_width(DASH, 7,  90),   # H
        col_width(DASH, 8,  90),   # I
        col_width(DASH, 9,  60),   # J
        col_width(DASH, 10, 100),  # K
        col_width(DASH, 11, 20),   # L spacer
        col_width(DASH, 12, 160),  # M
        col_width(DASH, 13, 70),   # N
        col_width(DASH, 14, 90),   # O
        col_width(DASH, 15, 90),   # P
        col_width(DASH, 16, 60),   # Q
        col_width(DASH, 17, 100),  # R
        col_width(DASH, 18, 20),   # S spacer
        col_width(DASH, 19, 140),  # T
        col_width(DASH, 20, 120),  # U
        col_width(DASH, 21, 120),  # V
        col_width(DASH, 22, 120),  # W

        # Freeze row 1
        freeze(DASH, rows=1),

        # Currency format — income Actual/Expected/Variance cols B–D rows 11–27
        number_fmt(DASH, 10, 1, 24, 4, '$#,##0.00'),
        # Needs H,I
        number_fmt(DASH, 10, 7, 24, 10, '$#,##0.00'),
        # Needs J (progress %)
        number_fmt(DASH, 10, 9, 24, 10, '0%'),
        # Wants O,P
        number_fmt(DASH, 10, 14, 24, 17, '$#,##0.00'),
        # Wants Q (progress %)
        number_fmt(DASH, 10, 16, 24, 17, '0%'),
        # Amount left — big font
        font(DASH, 10, 19, 11, 20, size=28, bold=True, colour=DARK_TEXT),
        number_fmt(DASH, 10, 19, 11, 20, '$#,##0.00'),
        # Percentage left
        number_fmt(DASH, 13, 19, 14, 20, '0.00%'),
        font(DASH, 13, 19, 14, 20, size=14, bold=True, colour=DARK_TEXT),
        # Status message
        font(DASH, 14, 19, 15, 20, size=10, bold=True, colour=DARK_TEXT),
    ]

    # Alt rows — Income table rows 11–23 (0-indexed 10–22)
    for i in range(12):
        colour = WHITE if i % 2 == 0 else ALT_ROW
        reqs.append(bg(DASH, 10 + i, 0, 11 + i, 4, colour))

    # Total income row
    reqs += [
        bg(DASH, 23, 0, 24, 4, LIGHT_SAGE),
        font(DASH, 23, 0, 24, 4, size=9, bold=True, colour=WHITE),
    ]
    # Rollover / Total Budget rows
    reqs += [
        bg(DASH, 25, 0, 26, 4, SOFT_MINT),
        bg(DASH, 26, 0, 27, 4, SAGE),
        font(DASH, 26, 0, 27, 4, size=9, bold=True, colour=WHITE),
    ]

    # Alt rows — Needs table
    for i in range(len(NEEDS_CATS)):
        colour = WHITE if i % 2 == 0 else ALT_ROW
        reqs.append(bg(DASH, 10 + i, 5, 11 + i, 11, colour))
    reqs += [
        bg(DASH, 10 + len(NEEDS_CATS), 5, 11 + len(NEEDS_CATS), 11, LIGHT_SAGE),
        font(DASH, 10 + len(NEEDS_CATS), 5, 11 + len(NEEDS_CATS), 11, size=9, bold=True, colour=WHITE),
    ]

    # Alt rows — Wants table
    for i in range(len(WANTS_CATS)):
        colour = WHITE if i % 2 == 0 else ALT_ROW
        reqs.append(bg(DASH, 10 + i, 12, 11 + i, 18, colour))
    reqs += [
        bg(DASH, 10 + len(WANTS_CATS), 12, 11 + len(WANTS_CATS), 18, LIGHT_SAGE),
        font(DASH, 10 + len(WANTS_CATS), 12, 11 + len(WANTS_CATS), 18, size=9, bold=True, colour=WHITE),
    ]

    # Borders — Income table
    reqs += [
        borders_outer(DASH, 9, 0, 24, 4),
        borders_inner(DASH, 9, 0, 24, 4),
    ]
    # Borders — Needs table
    reqs += [
        borders_outer(DASH, 9, 5, 10 + len(NEEDS_CATS) + 1, 11),
        borders_inner(DASH, 9, 5, 10 + len(NEEDS_CATS) + 1, 11),
    ]
    # Borders — Wants table
    reqs += [
        borders_outer(DASH, 9, 12, 10 + len(WANTS_CATS) + 1, 18),
        borders_inner(DASH, 9, 12, 10 + len(WANTS_CATS) + 1, 18),
    ]

    # Month selector validation
    reqs.append({
        "setDataValidation": {
            "range": grid(DASH, 4, 1, 5, 2),
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": m} for m in MONTH_OPTIONS],
                },
                "showCustomUi": True,
                "strict": True,
            },
        }
    })

    batch(service, sid, reqs)


# ── Dashboard conditional formatting ─────────────────────────────────────────

def setup_dashboard_cf(service, sid):
    """Conditional formatting rules for Dashboard."""
    # Needs Actual col I (0-indexed col 8), rows 11–22 (0-indexed 10–21)
    # Wants Actual col P (0-indexed col 15)
    actual_ranges = [
        grid(DASH, 10, 8, 22, 9),   # Needs actual (I11:I22)
        grid(DASH, 10, 15, 22, 16), # Wants actual (P11:P22)
    ]

    # Over budget (actual > expected)
    for rng in actual_ranges:
        # For needs: expected is col H (7), actual col I (8)
        # We use a CUSTOM_FORMULA per range
        pass

    reqs = []

    # Needs: over budget — I > H
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 8, 22, 9)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=AND(H11<>0,I11>H11)"}],
                    },
                    "format": {"backgroundColor": RED_BG, "textFormat": {"foregroundColor": RED_TEXT}},
                },
            },
            "index": 0,
        }
    })
    # Needs: on track
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 8, 22, 9)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=AND(H11<>0,I11<=H11)"}],
                    },
                    "format": {"backgroundColor": GREEN_BG, "textFormat": {"foregroundColor": GREEN_TEXT}},
                },
            },
            "index": 1,
        }
    })

    # Wants: over budget — P > O
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 15, 22, 16)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=AND(O11<>0,P11>O11)"}],
                    },
                    "format": {"backgroundColor": RED_BG, "textFormat": {"foregroundColor": RED_TEXT}},
                },
            },
            "index": 2,
        }
    })
    # Wants: on track
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 15, 22, 16)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=AND(O11<>0,P11<=O11)"}],
                    },
                    "format": {"backgroundColor": GREEN_BG, "textFormat": {"foregroundColor": GREEN_TEXT}},
                },
            },
            "index": 3,
        }
    })

    # Progress % — Needs col J (9), over 100%
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 9, 22, 10)],
                "booleanRule": {
                    "condition": {
                        "type": "NUMBER_GREATER",
                        "values": [{"userEnteredValue": "1"}],
                    },
                    "format": {"backgroundColor": RED_BG},
                },
            },
            "index": 4,
        }
    })
    # 75–100%
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 9, 22, 10)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=AND(J11>=0.75,J11<=1)"}],
                    },
                    "format": {"backgroundColor": YELLOW_BG},
                },
            },
            "index": 5,
        }
    })
    # Under 75%
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 10, 9, 22, 10)],
                "booleanRule": {
                    "condition": {
                        "type": "NUMBER_LESS",
                        "values": [{"userEnteredValue": "0.75"}],
                    },
                    "format": {"backgroundColor": GREEN_BG},
                },
            },
            "index": 6,
        }
    })

    # KPI totals vs targets (row 6, 0-indexed)
    # Needs total C7 (col 2) vs T2 (Needs target)
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 6, 2, 7, 3)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=C7>$T$2"}],
                    },
                    "format": {"backgroundColor": RED_BG, "textFormat": {"foregroundColor": RED_TEXT}},
                },
            },
            "index": 7,
        }
    })
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 6, 5, 7, 6)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=F7>$U$2"}],
                    },
                    "format": {"backgroundColor": RED_BG, "textFormat": {"foregroundColor": RED_TEXT}},
                },
            },
            "index": 8,
        }
    })
    reqs.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid(DASH, 6, 8, 7, 9)],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=I7>$V$2"}],
                    },
                    "format": {"backgroundColor": RED_BG, "textFormat": {"foregroundColor": RED_TEXT}},
                },
            },
            "index": 9,
        }
    })

    batch(service, sid, reqs)


# ── Charts ────────────────────────────────────────────────────────────────────

def add_charts(service, sid):
    reqs = []

    # Chart 1: Cash Flow Summary (Bar) — Dashboard, anchored col A row 10 overlay
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "CASH FLOW SUMMARY",
                    "basicChart": {
                        "chartType": "BAR",
                        "legendPosition": "TOP_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS"},
                            {"position": "LEFT_AXIS"},
                        ],
                        "domains": [{
                            "domain": {
                                "sourceRange": {
                                    "sources": [{"sheetId": DASH,
                                                 "startRowIndex": 10, "endRowIndex": 22,
                                                 "startColumnIndex": 0, "endColumnIndex": 1}]
                                }
                            }
                        }],
                        "series": [
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{"sheetId": DASH,
                                                     "startRowIndex": 10, "endRowIndex": 22,
                                                     "startColumnIndex": 1, "endColumnIndex": 2}]
                                    }
                                },
                                "targetAxis": "BOTTOM_AXIS",
                                "color": c("A8C5B0"),
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [{"sheetId": DASH,
                                                     "startRowIndex": 10, "endRowIndex": 22,
                                                     "startColumnIndex": 2, "endColumnIndex": 3}]
                                    }
                                },
                                "targetAxis": "BOTTOM_AXIS",
                                "color": SAGE,
                            },
                        ],
                        "headerCount": 0,
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": DASH, "rowIndex": 28, "columnIndex": 0},
                        "widthPixels": 380,
                        "heightPixels": 260,
                    }
                },
            }
        }
    })

    # Chart 2: Donut — Actual Allocation
    reqs.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": "ACTUAL ALLOCATION SUMMARY",
                    "pieChart": {
                        "legendPosition": "RIGHT_LEGEND",
                        "pieHole": 0.5,
                        "domain": {
                            "sourceRange": {
                                "sources": [{"sheetId": DASH,
                                             "startRowIndex": 6, "endRowIndex": 7,
                                             "startColumnIndex": 1, "endColumnIndex": 9}]
                            }
                        },
                        "series": {
                            "sourceRange": {
                                "sources": [{"sheetId": DASH,
                                             "startRowIndex": 6, "endRowIndex": 7,
                                             "startColumnIndex": 2, "endColumnIndex": 9}]
                            }
                        },
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": DASH, "rowIndex": 28, "columnIndex": 5},
                        "widthPixels": 300,
                        "heightPixels": 260,
                    }
                },
            }
        }
    })

    batch(service, sid, reqs)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Authenticating…")
    service = get_service()

    print("Creating spreadsheet…")
    sid = create_spreadsheet(service)
    print(f"  Created: https://docs.google.com/spreadsheets/d/{sid}")

    print("Setting up Reference tab…")
    setup_reference(service, sid)

    print("Setting up Settings tab…")
    setup_settings(service, sid)

    print("Setting up Transactions Log…")
    setup_transactions_log(service, sid)

    print("Writing Dashboard values & formulas…")
    setup_dashboard_values(service, sid)

    print("Formatting Dashboard…")
    setup_dashboard_format(service, sid)

    print("Applying Dashboard conditional formatting…")
    setup_dashboard_cf(service, sid)

    print("Adding charts…")
    add_charts(service, sid)

    print("\n✓ Done!")
    print(f"  Open: https://docs.google.com/spreadsheets/d/{sid}/edit")
    return sid


if __name__ == "__main__":
    main()
