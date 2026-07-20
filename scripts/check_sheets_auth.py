"""Verify the Google Sheets service-account credential end to end.

Run this ONCE after dropping the service-account JSON key into the path below
and sharing the Finance sheet with the service account's email. It proves all
three things that can independently be wrong — the key parses, the Sheets API
is enabled on the project, and the sheet has actually been shared with the
service account — and says WHICH one failed rather than a bare stack trace.

    python scripts/check_sheets_auth.py

It only READS (opens the spreadsheet and lists tab names). Nothing is written,
so it is safe to run against the live Finance sheet at any time.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CFG  # noqa: E402
from ssl_certs import ensure_ca_bundle  # noqa: E402

ensure_ca_bundle()

# Windows' console codepage (cp1252) can't encode the em dash used below —
# same cosmetic-print crash guarded in verify_pipeline.py and spending_summary.py.
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Outside the repo on purpose: the repo lives in OneDrive, and a private key
# must not sync to the cloud. Overridable via config.json -> sheetsServiceAccountKey.
DEFAULT_KEY = os.path.join(os.path.expanduser("~"), ".secrets",
                           "finance-sheets-sync.json")
KEY_PATH = os.path.expanduser(CFG.get("sheetsServiceAccountKey") or DEFAULT_KEY)
SHEET_ID = CFG["financeSheetId"]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def fail(what, detail, fix):
    print(f"\nFAILED: {what}\n  {detail}\n  Fix: {fix}")
    sys.exit(1)


def main():
    if not os.path.exists(KEY_PATH):
        fail("no service-account key found",
             f"expected it at {KEY_PATH}",
             "download the JSON key from the service account's Keys tab and save it there")

    try:
        with open(KEY_PATH, encoding="utf-8") as f:
            info = json.load(f)
    except (OSError, ValueError) as e:
        fail("the key file is not readable JSON", str(e),
             "re-download it — the JSON key, not the P12")

    if info.get("type") != "service_account":
        fail("that is not a service-account key",
             f"its 'type' is {info.get('type')!r}",
             "create the key from the SERVICE ACCOUNT's Keys tab, not from OAuth client IDs")

    email = info.get("client_email", "(none)")
    print(f"Key    : {KEY_PATH}")
    print(f"Account: {email}")
    print(f"Sheet  : {SHEET_ID}")

    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    client = gspread.authorize(creds)

    try:
        sh = client.open_by_key(SHEET_ID)
    except Exception as e:                                   # noqa: BLE001
        msg = str(e)
        if "403" in msg and "disabled" in msg.lower():
            fail("the Sheets API is not enabled on this project", msg,
                 "enable it at console.cloud.google.com/apis/library/sheets.googleapis.com")
        if "403" in msg or "PERMISSION_DENIED" in msg:
            fail("the sheet has not been shared with the service account", msg,
                 f"open the Finance sheet -> Share -> add {email} as Editor")
        if "404" in msg:
            fail("no sheet with that id", msg,
                 "check financeSheetId in scripts/config.json")
        fail("could not open the spreadsheet", msg, "see the message above")

    tabs = [ws.title for ws in sh.worksheets()]
    print(f"\nOK — opened '{sh.title}' with {len(tabs)} tab(s):")
    for t in tabs:
        print(f"  - {t}")
    print("\nAuth is working. Read-only check; nothing was written.")


if __name__ == "__main__":
    main()
