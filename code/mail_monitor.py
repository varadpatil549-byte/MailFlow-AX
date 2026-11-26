import time
from pathlib import Path
from datetime import datetime

import pythoncom
import win32com.client
from openpyxl import load_workbook

from ai_classifier import classify_email


# =========================================================
# CONFIGURATION
# =========================================================

PROJECT_ROOT = Path(__file__).resolve().parent

EXCEL_FILE = (
    PROJECT_ROOT
    / "data"
    / "MailFlow_AX.xlsx"
)

INBOX_FOLDER = 6
MAIL_ITEM_CLASS = 43

POLL_INTERVAL = 2
CHECK_LATEST = 50


# =========================================================
# OUTLOOK CONNECTION
# =========================================================

def connect_outlook():
    """Connect to the already-running Classic Outlook."""

    try:
        outlook = win32com.client.GetActiveObject(
            "Outlook.Application"
        )

        print("✅ Connected to running Outlook")

    except Exception:

        outlook = win32com.client.Dispatch(
            "Outlook.Application"
        )

        print("✅ Outlook started")

    return outlook.GetNamespace("MAPI")


# =========================================================
# EMAIL FUNCTIONS
# =========================================================

def is_mail(item):
    """Check whether the Outlook item is an email."""

    return (
        getattr(item, "Class", None)
        == MAIL_ITEM_CLASS
    )


def extract_email(item):
    """Extract the information we need from Outlook."""

    try:
        sender_email = (
            item.SenderEmailAddress or ""
        )
    except Exception:
        sender_email = ""

    try:
        body = item.Body or ""

        body = (
            body
            .replace("\r", " ")
            .replace("\n", " ")
            .strip()
        )

    except Exception:
        body = ""

    return {
        "entry_id": item.EntryID,

        "subject": (
            item.Subject
            or "(No subject)"
        ),

        "sender_name": (
            item.SenderName
            or "(Unknown sender)"
        ),

        "sender_email": sender_email,

        "received_time": item.ReceivedTime,

        "body": body[:5000],
    }


# =========================================================
# EXCEL
# =========================================================

def save_to_excel(
    email,
    ai_result,
):
    """Write the processed email into Email_Log."""

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE
    )

    try:

        sheet = workbook["Email_Log"]

        # Header is row 1.
        # Therefore first email gets MFAX-00001.
        ticket_number = (
            sheet.max_row
        )

        ticket_id = (
            f"MFAX-{ticket_number:05d}"
        )

        captured_time = (
            datetime.now()
            .strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )

        sheet.append([
            ticket_id,
            email["entry_id"],
            str(email["received_time"]),
            email["sender_name"],
            email["sender_email"],
            email["subject"],
            email["body"],
            ai_result["category"],
            ai_result["confidence"],
            captured_time,
            "Captured",
        ])

        workbook.save(
            EXCEL_FILE
        )

        return ticket_id

    finally:

        workbook.close()


# =========================================================
# MAIN MONITOR
# =========================================================

def main():

    print("\n" + "=" * 70)
    print("MAILFLOW AX - PHASE 1")
    print("OUTLOOK → LOCAL AI → EXCEL")
    print("=" * 70)

    print(
        f"\nExcel file:\n{EXCEL_FILE}"
    )

    if not EXCEL_FILE.exists():

        print(
            "\n❌ Excel file does not exist."
        )

        return

    pythoncom.CoInitialize()

    try:

        namespace = connect_outlook()

        inbox = (
            namespace
            .GetDefaultFolder(
                INBOX_FOLDER
            )
        )

        print(
            f"✅ Monitoring folder: "
            f"{inbox.Name}"
        )

        # -------------------------------------------------
        # Existing emails
        #
        # These are intentionally ignored.
        # -------------------------------------------------

        seen_ids = set()

        items = inbox.Items

        items.Sort(
            "[ReceivedTime]",
            True
        )

        existing_count = 0

        for item in items:

            if not is_mail(item):
                continue

            entry_id = getattr(
                item,
                "EntryID",
                None
            )

            if entry_id:

                seen_ids.add(
                    entry_id
                )

            existing_count += 1

            if (
                existing_count
                >= CHECK_LATEST
            ):
                break

        print(
            f"✅ Existing emails ignored: "
            f"{existing_count}"
        )

        print("\n" + "=" * 70)
        print("🟢 MAILFLOW AX IS RUNNING")
        print(
            "Waiting for NEW incoming emails..."
        )
        print(
            "Outlook → Qwen3 → Excel"
        )
        print(
            f"Polling every "
            f"{POLL_INTERVAL} seconds"
        )
        print("=" * 70)

        # -------------------------------------------------
        # Continuous monitoring
        # -------------------------------------------------

        while True:

            try:

                items = inbox.Items

                items.Sort(
                    "[ReceivedTime]",
                    True
                )

                checked = 0

                for item in items:

                    if (
                        checked
                        >= CHECK_LATEST
                    ):
                        break

                    if not is_mail(item):
                        continue

                    checked += 1

                    entry_id = getattr(
                        item,
                        "EntryID",
                        None
                    )

                    if not entry_id:
                        continue

                    # Already processed
                    if entry_id in seen_ids:
                        continue

                    # -------------------------------------------------
                    # NEW EMAIL DETECTED
                    # -------------------------------------------------

                    email = extract_email(
                        item
                    )

                    print(
                        "\n" + "=" * 70
                    )

                    print(
                        "📨 NEW EMAIL DETECTED"
                    )

                    print(
                        "=" * 70
                    )

                    print(
                        f"Subject  : "
                        f"{email['subject']}"
                    )

                    print(
                        f"Sender   : "
                        f"{email['sender_email']}"
                    )

                    print(
                        f"Received : "
                        f"{email['received_time']}"
                    )

                    # -------------------------------------------------
                    # AI CLASSIFICATION
                    # -------------------------------------------------

                    print(
                        "\n🤖 AI classifying email..."
                    )

                    try:

                        ai_result = (
                            classify_email(
                                email["subject"],
                                email["body"],
                            )
                        )

                        print(
                            f"✅ Category   : "
                            f"{ai_result['category']}"
                        )

                        print(
                            f"✅ Confidence : "
                            f"{ai_result['confidence']}%"
                        )

                        print(
                            f"ℹ️ Reason     : "
                            f"{ai_result['reason']}"
                        )

                    except Exception as error:

                        print(
                            "\n❌ AI classification "
                            "failed:"
                        )

                        print(
                            error
                        )

                        print(
                            "Email will be retried."
                        )

                        continue

                    # -------------------------------------------------
                    # EXCEL
                    # -------------------------------------------------

                    print(
                        "\n📝 Writing to Excel..."
                    )

                    try:

                        ticket_id = (
                            save_to_excel(
                                email,
                                ai_result,
                            )
                        )

                        # Mark as processed ONLY
                        # after Excel succeeds.
                        seen_ids.add(
                            entry_id
                        )

                        print(
                            "✅ Excel updated successfully"
                        )

                        print(
                            f"   Ticket     : "
                            f"{ticket_id}"
                        )

                        print(
                            f"   Category   : "
                            f"{ai_result['category']}"
                        )

                        print(
                            f"   Confidence : "
                            f"{ai_result['confidence']}%"
                        )

                        print(
                            "\n🎯 COMPLETE"
                        )

                        print(
                            "Outlook → AI → Excel"
                        )

                    except PermissionError:

                        print(
                            "\n❌ Excel is open/locked."
                        )

                        print(
                            "Close "
                            "MailFlow_AX.xlsx "
                            "and retry."
                        )

                    except Exception as error:

                        print(
                            "\n❌ Excel logging failed:"
                        )

                        print(
                            error
                        )

                    print(
                        "=" * 70
                    )

            except Exception as error:

                print(
                    f"\n⚠️ Monitor error: "
                    f"{error}"
                )

            time.sleep(
                POLL_INTERVAL
            )

    except KeyboardInterrupt:

        print(
            "\n🛑 MailFlow AX stopped."
        )

    finally:

        pythoncom.CoUninitialize()


if __name__ == "__main__":

    main()