import time
from pathlib import Path
from datetime import datetime

import pythoncom
import win32com.client
from openpyxl import load_workbook


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

# Round-robin users.
USERS = [
    "User1",
    "User2",
    "User3",
]


# =========================================================
# OUTLOOK
# =========================================================

def connect_outlook():
    """Connect to Classic Outlook."""

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


def is_mail(item):
    """Check whether the Outlook item is a mail item."""

    return (
        getattr(item, "Class", None)
        == MAIL_ITEM_CLASS
    )


# =========================================================
# EXCEL ALLOCATION LOG
# =========================================================

def prepare_allocation_log():

    workbook = load_workbook(
        EXCEL_FILE
    )

    try:

        if "Allocation_Log" not in workbook.sheetnames:

            sheet = workbook.create_sheet(
                "Allocation_Log"
            )

            sheet.append([
                "Allocation_ID",
                "Message_ID",
                "Subject",
                "Sender",
                "Assigned_User",
                "Destination_Folder",
                "Allocation_Time",
                "Status",
            ])

            workbook.save(
                EXCEL_FILE
            )

            print(
                "✅ Allocation_Log sheet created"
            )

    finally:

        workbook.close()


def log_allocation(
    message_id,
    subject,
    sender,
    assigned_user,
    destination_folder,
    status,
):

    workbook = load_workbook(
        EXCEL_FILE
    )

    try:

        sheet = workbook[
            "Allocation_Log"
        ]

        allocation_number = (
            sheet.max_row
        )

        allocation_id = (
            f"ALLOC-{allocation_number:05d}"
        )

        allocation_time = (
            datetime.now()
            .strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )

        sheet.append([
            allocation_id,
            message_id,
            subject,
            sender,
            assigned_user,
            destination_folder,
            allocation_time,
            status,
        ])

        workbook.save(
            EXCEL_FILE
        )

        return allocation_id

    finally:

        workbook.close()


# =========================================================
# MAIN
# =========================================================

def main():

    print("\n" + "=" * 70)
    print("MAILFLOW AX - PHASE 2")
    print("ROUND-ROBIN EMAIL ALLOCATION")
    print("=" * 70)

    if not EXCEL_FILE.exists():

        print(
            f"❌ Excel file not found:"
            f"\n{EXCEL_FILE}"
        )

        return

    # Prepare Excel before starting.
    prepare_allocation_log()

    pythoncom.CoInitialize()

    try:

        namespace = connect_outlook()

        inbox = (
            namespace.GetDefaultFolder(
                INBOX_FOLDER
            )
        )

        print(
            f"✅ Inbox found: "
            f"{inbox.Name}"
        )

        # -------------------------------------------------
        # Find User1 / User2 / User3 folders
        # -------------------------------------------------

        destination_folders = {}

        for user in USERS:

            try:

                folder = inbox.Folders[user]

                destination_folders[user] = folder

                print(
                    f"✅ Folder found: "
                    f"{user}"
                )

            except Exception:

                print(
                    f"❌ Folder not found: "
                    f"{user}"
                )

                return

        # -------------------------------------------------
        # Existing emails are ignored
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

            if existing_count >= CHECK_LATEST:
                break

        # -------------------------------------------------
        # Round robin starts from User1
        # -------------------------------------------------

        round_robin_index = 0

        print(
            f"\n✅ Existing emails ignored: "
            f"{existing_count}"
        )

        print(
            "\nRound-robin sequence:"
        )

        print(
            " → ".join(USERS)
        )

        print("\n" + "=" * 70)

        print(
            "🟢 MAILFLOW AX PHASE 2 IS RUNNING"
        )

        print(
            "Waiting for NEW emails..."
        )

        print(
            "New emails will be MOVED automatically."
        )

        print(
            "Press Ctrl+C to stop."
        )

        print("=" * 70)

        # -------------------------------------------------
        # MONITOR
        # -------------------------------------------------

        while True:

            try:

                items = inbox.Items

                items.Sort(
                    "[ReceivedTime]",
                    True
                )

                checked = 0

                # We scan newest first.
                # Only a small number of items
                # need to be inspected.
                for item in items:

                    if checked >= CHECK_LATEST:
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

                    # Existing / already processed
                    if entry_id in seen_ids:
                        continue

                    # -------------------------------------------------
                    # NEW EMAIL
                    # -------------------------------------------------

                    subject = (
                        item.Subject
                        or "(No subject)"
                    )

                    sender = (
                        item.SenderName
                        or "(Unknown)"
                    )

                    assigned_user = USERS[
                        round_robin_index
                    ]

                    destination = (
                        destination_folders[
                            assigned_user
                        ]
                    )

                    print(
                        "\n" + "=" * 70
                    )

                    print(
                        "📨 NEW EMAIL"
                    )

                    print(
                        f"Subject : "
                        f"{subject}"
                    )

                    print(
                        f"Sender  : "
                        f"{sender}"
                    )

                    print(
                        f"➡️ Assigned to: "
                        f"{assigned_user}"
                    )

                    print(
                        f"📁 Moving to: "
                        f"{destination.Name}"
                    )

                    # -------------------------------------------------
                    # MOVE
                    # -------------------------------------------------

                    try:

                        item.Move(
                            destination
                        )

                        # Log AFTER successful move.
                        allocation_id = (
                            log_allocation(
                                message_id=entry_id,
                                subject=subject,
                                sender=sender,
                                assigned_user=assigned_user,
                                destination_folder=(
                                    destination.Name
                                ),
                                status="Moved",
                            )
                        )

                        print(
                            "✅ EMAIL MOVED"
                        )

                        print(
                            f"✅ Allocation ID: "
                            f"{allocation_id}"
                        )

                        print(
                            f"✅ Destination: "
                            f"{destination.Name}"
                        )

                        # Advance round robin.
                        round_robin_index = (
                            round_robin_index + 1
                        ) % len(USERS)

                        # The email was successfully
                        # handled.
                        seen_ids.add(
                            entry_id
                        )

                    except PermissionError:

                        print(
                            "❌ Excel is locked."
                        )

                    except Exception as error:

                        print(
                            "❌ Email move failed:"
                        )

                        print(
                            error
                        )

                    print(
                        "=" * 70
                    )

            except Exception as error:

                print(
                    f"\n⚠️ Monitor error:"
                )

                print(error)

            time.sleep(
                POLL_INTERVAL
            )

    except KeyboardInterrupt:

        print(
            "\n🛑 Phase 2 stopped."
        )

    finally:

        pythoncom.CoUninitialize()


if __name__ == "__main__":

    main()