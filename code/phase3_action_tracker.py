import time
from pathlib import Path
from datetime import datetime

import pythoncom
import win32com.client
from openpyxl import load_workbook


# =========================================================
# MAILFLOW AX - PHASE 3
# USER ACTION TRACKER
#
# Actioned = YES ONLY when:
#   - Reply is sent
#   - Reply All is sent
#   - Forward is sent
#
# Category / flag changes are logged separately, but they
# DO NOT make the email "Actioned".
# =========================================================

PROJECT_ROOT = Path(__file__).resolve().parent
EXCEL_FILE = PROJECT_ROOT / "data" / "MailFlow_AX.xlsx"

INBOX_FOLDER = 6
SENT_ITEMS_FOLDER = 5
MAIL_ITEM_CLASS = 43

USER_FOLDERS = [
    "User1",
    "User2",
    "User3",
]

POLL_INTERVAL = 3
MAX_MAIL_SCAN = 100


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def connect_outlook():
    """Connect to the running Classic Outlook application."""
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        print("✅ Connected to running Outlook")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("✅ Outlook started")

    return outlook.GetNamespace("MAPI")


def is_mail(item):
    """Return True only for Outlook MailItem objects."""
    return getattr(item, "Class", None) == MAIL_ITEM_CLASS


def normalize_subject(subject):
    """
    Normalize an Outlook subject so RE:/FW:/FWD: versions
    can be compared with the original subject.
    """
    if not subject:
        return ""

    value = str(subject).strip()

    changed = True
    while changed:
        changed = False
        for prefix in ("re:", "fw:", "fwd:"):
            if value.lower().startswith(prefix):
                value = value[len(prefix):].strip()
                changed = True

    for prefix in ("[external]", "[extern]"):
        if value.lower().startswith(prefix):
            value = value[len(prefix):].strip()

    return value.lower().strip()


def is_reply_or_forward_subject(subject):
    """Return the business action represented by a sent subject."""
    if not subject:
        return None

    value = subject.strip().lower()

    if value.startswith("re:"):
        return "Reply / Reply All"

    if value.startswith("fw:") or value.startswith("fwd:"):
        return "Forward"

    return None


ACTION_HEADERS = [
    "Message_ID",
    "Assigned_Folder",
    "Subject",
    "Received_Time",
    "Actioned",
    "Final_Action_Type",
    "Final_Action_Time",
    "Reply_Forward_Subject",
    "Category_Changed",
    "Category_Change_Time",
    "Current_Categories",
    "Flag_Changed",
    "Flag_Change_Time",
    "Current_Flag",
    "Last_Checked",
]


def prepare_action_log():
    """Create Action_Log if it doesn't exist."""
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(EXCEL_FILE)

    try:
        if "Action_Log" not in workbook.sheetnames:
            sheet = workbook.create_sheet("Action_Log")
            sheet.append(ACTION_HEADERS)
            workbook.save(EXCEL_FILE)
            print("✅ Action_Log sheet created")
        else:
            sheet = workbook["Action_Log"]
            existing_headers = [
                sheet.cell(1, col).value
                for col in range(1, len(ACTION_HEADERS) + 1)
            ]
            if existing_headers != ACTION_HEADERS:
                raise ValueError(
                    "Action_Log already exists with a different "
                    "header structure. Remove the old Action_Log "
                    "sheet once for this Phase 3 test."
                )
    finally:
        workbook.close()


def find_action_row(sheet, message_id):
    """Find the Action_Log row for a message."""
    for row in range(2, sheet.max_row + 1):
        if str(sheet.cell(row, 1).value or "") == str(message_id):
            return row
    return None


def ensure_action_row(record):
    """Create the Action_Log row if it does not exist."""
    workbook = load_workbook(EXCEL_FILE)

    try:
        sheet = workbook["Action_Log"]
        row = find_action_row(sheet, record["message_id"])

        if row is None:
            sheet.append([
                record["message_id"],
                record["assigned_folder"],
                record["subject"],
                str(record["received_time"]),
                "NO",
                "",
                "",
                "",
                "NO",
                "",
                record["initial_categories"],
                "NO",
                "",
                str(record["initial_flag"]),
                now_text(),
            ])
            workbook.save(EXCEL_FILE)
    finally:
        workbook.close()


def update_action_row(record):
    """Update the existing Action_Log row."""
    workbook = load_workbook(EXCEL_FILE)

    try:
        sheet = workbook["Action_Log"]
        row = find_action_row(sheet, record["message_id"])

        if row is None:
            # Write the row directly here; avoids recursion.
            sheet.append([
                record["message_id"],
                record["assigned_folder"],
                record["subject"],
                str(record["received_time"]),
                "YES" if record["actioned"] else "NO",
                record["final_action_type"],
                record["final_action_time"],
                record["reply_forward_subject"],
                "YES" if record["category_changed"] else "NO",
                record["category_change_time"],
                record["current_categories"],
                "YES" if record["flag_changed"] else "NO",
                record["flag_change_time"],
                str(record["current_flag"]),
                now_text(),
            ])
        else:
            values = [
                record["message_id"],
                record["assigned_folder"],
                record["subject"],
                str(record["received_time"]),
                "YES" if record["actioned"] else "NO",
                record["final_action_type"],
                record["final_action_time"],
                record["reply_forward_subject"],
                "YES" if record["category_changed"] else "NO",
                record["category_change_time"],
                record["current_categories"],
                "YES" if record["flag_changed"] else "NO",
                record["flag_change_time"],
                str(record["current_flag"]),
                now_text(),
            ]

            for col, value in enumerate(values, start=1):
                sheet.cell(row, col).value = value

        workbook.save(EXCEL_FILE)
    finally:
        workbook.close()


def get_allocated_emails(namespace):
    """
    Read emails currently sitting in User1/User2/User3
    and create an in-memory tracking state.
    """
    inbox = namespace.GetDefaultFolder(INBOX_FOLDER)
    tracked = {}

    for folder_name in USER_FOLDERS:
        try:
            folder = inbox.Folders[folder_name]
        except Exception:
            print(f"⚠️ Allocation folder not found: {folder_name}")
            continue

        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        scanned = 0

        for item in items:
            if scanned >= MAX_MAIL_SCAN:
                break
            if not is_mail(item):
                continue

            scanned += 1

            try:
                message_id = item.EntryID
                if not message_id:
                    continue

                try:
                    conversation_id = item.ConversationID or ""
                except Exception:
                    conversation_id = ""

                try:
                    conversation_topic = item.ConversationTopic or ""
                except Exception:
                    conversation_topic = ""

                categories = item.Categories or ""
                flag_status = getattr(item, "FlagStatus", None)

                record = {
                    "message_id": message_id,
                    "assigned_folder": folder_name,
                    "subject": item.Subject or "(No subject)",
                    "received_time": item.ReceivedTime,
                    "conversation_id": conversation_id,
                    "conversation_topic": conversation_topic,
                    "normalized_subject": normalize_subject(item.Subject or ""),
                    "initial_categories": categories,
                    "current_categories": categories,
                    "initial_flag": flag_status,
                    "current_flag": flag_status,
                    "category_changed": False,
                    "category_change_time": "",
                    "flag_changed": False,
                    "flag_change_time": "",
                    "actioned": False,
                    "final_action_type": "",
                    "final_action_time": "",
                    "reply_forward_subject": "",
                }

                tracked[message_id] = record
                ensure_action_row(record)
            except Exception:
                continue

    return tracked


def get_existing_sent_ids(namespace):
    """
    Capture current Sent Items. Only NEW sent items after
    tracker startup are considered for reply/forward detection.
    """
    sent_items = namespace.GetDefaultFolder(SENT_ITEMS_FOLDER)
    existing_ids = set()

    items = sent_items.Items
    items.Sort("[SentOn]", True)

    scanned = 0

    for item in items:
        if scanned >= MAX_MAIL_SCAN:
            break
        if not is_mail(item):
            continue

        scanned += 1

        try:
            entry_id = item.EntryID
            if entry_id:
                existing_ids.add(entry_id)
        except Exception:
            continue

    return existing_ids


def check_category_and_flag_changes(namespace, tracked):
    """
    Category and flag changes are logged but never make
    Actioned = YES.
    """
    inbox = namespace.GetDefaultFolder(INBOX_FOLDER)

    for message_id, record in tracked.items():
        try:
            folder = inbox.Folders[record["assigned_folder"]]
            found_item = None

            for item in folder.Items:
                if not is_mail(item):
                    continue
                try:
                    if item.EntryID == message_id:
                        found_item = item
                        break
                except Exception:
                    continue

            if found_item is None:
                continue

            current_categories = found_item.Categories or ""

            if current_categories != record["current_categories"]:
                record["current_categories"] = current_categories
                record["category_changed"] = True

                if not record["category_change_time"]:
                    record["category_change_time"] = now_text()

                print("\n" + "-" * 65)
                print("🏷️ CATEGORY CHANGE DETECTED")
                print(f"Subject  : {record['subject']}")
                print(f"Folder   : {record['assigned_folder']}")
                print(f"Category : {current_categories or '(None)'}")
                print("Actioned remains: NO")
                print("-" * 65)

                update_action_row(record)

            current_flag = getattr(found_item, "FlagStatus", None)

            if current_flag != record["current_flag"]:
                record["current_flag"] = current_flag
                record["flag_changed"] = True

                if not record["flag_change_time"]:
                    record["flag_change_time"] = now_text()

                print("\n" + "-" * 65)
                print("🚩 FLAG CHANGE DETECTED")
                print(f"Subject : {record['subject']}")
                print(f"Folder  : {record['assigned_folder']}")
                print(f"Flag    : {current_flag}")
                print("Actioned remains: NO")
                print("-" * 65)

                update_action_row(record)

        except Exception:
            continue


def find_matching_record(sent_item, tracked):
    """
    Match a newly sent Reply/Reply All/Forward to an allocated
    message using ConversationID, ConversationTopic, then subject.
    """
    sent_subject = (sent_item.Subject or "").strip()
    action_type = is_reply_or_forward_subject(sent_subject)

    if not action_type:
        return None, None

    sent_normalized_subject = normalize_subject(sent_subject)

    try:
        sent_conversation_id = sent_item.ConversationID or ""
    except Exception:
        sent_conversation_id = ""

    try:
        sent_conversation_topic = sent_item.ConversationTopic or ""
    except Exception:
        sent_conversation_topic = ""

    # 1. Strongest match: ConversationID.
    for record in tracked.values():
        if record["actioned"]:
            continue

        if (
            sent_conversation_id
            and record["conversation_id"]
            and sent_conversation_id == record["conversation_id"]
        ):
            return record, action_type

    # 2. ConversationTopic.
    for record in tracked.values():
        if record["actioned"]:
            continue

        if (
            sent_conversation_topic
            and record["conversation_topic"]
            and sent_conversation_topic.lower()
            == record["conversation_topic"].lower()
        ):
            return record, action_type

    # 3. Normalized subject.
    matches = []

    for record in tracked.values():
        if record["actioned"]:
            continue

        if sent_normalized_subject == record["normalized_subject"]:
            matches.append(record)

    if len(matches) == 1:
        return matches[0], action_type

    return None, None


def check_new_sent_items(namespace, tracked, seen_sent_ids):
    """
    Only a new RE:/FW:/FWD: item in Sent Items can set
    Actioned = YES.
    """
    sent_items = namespace.GetDefaultFolder(SENT_ITEMS_FOLDER)
    items = sent_items.Items
    items.Sort("[SentOn]", True)

    scanned = 0

    for sent_item in items:
        if scanned >= MAX_MAIL_SCAN:
            break
        if not is_mail(sent_item):
            continue

        scanned += 1

        try:
            sent_id = sent_item.EntryID
            if not sent_id:
                continue

            if sent_id in seen_sent_ids:
                continue

            seen_sent_ids.add(sent_id)

            record, action_type = find_matching_record(
                sent_item,
                tracked,
            )

            if record is None:
                continue

            record["actioned"] = True
            record["final_action_type"] = action_type
            record["final_action_time"] = now_text()
            record["reply_forward_subject"] = sent_item.Subject or ""

            update_action_row(record)

            print("\n" + "=" * 70)
            print("✅ FINAL USER ACTION DETECTED")
            print("=" * 70)
            print(f"Original : {record['subject']}")
            print(f"Folder   : {record['assigned_folder']}")
            print(f"Action   : {action_type}")
            print(f"Sent Subj: {record['reply_forward_subject']}")
            print("Actioned : YES")
            print(f"Time     : {record['final_action_time']}")
            print("=" * 70)

        except Exception:
            continue


def main():
    print("\n" + "=" * 70)
    print("MAILFLOW AX - PHASE 3")
    print("REPLY / REPLY ALL / FORWARD ACTION TRACKER")
    print("=" * 70)

    if not EXCEL_FILE.exists():
        print(f"❌ Excel file not found:\n{EXCEL_FILE}")
        return

    prepare_action_log()

    pythoncom.CoInitialize()

    try:
        namespace = connect_outlook()

        tracked = get_allocated_emails(namespace)

        print(
            f"✅ Allocated emails tracked: {len(tracked)}"
        )

        if not tracked:
            print("\n⚠️ No emails found in User1/User2/User3.")
            print(
                "Run Phase 2 first and allocate at least "
                "one test email."
            )
            return

        seen_sent_ids = get_existing_sent_ids(namespace)

        print(
            f"✅ Existing Sent Items ignored: {len(seen_sent_ids)}"
        )

        print("\nTracked emails:")
        for record in tracked.values():
            print(
                f"  {record['assigned_folder']} "
                f"→ {record['subject']}"
            )

        print("\n" + "=" * 70)
        print("🟢 PHASE 3 IS RUNNING")
        print("Reply / Reply All / Forward = ACTIONED")
        print("Category / Flag = LOGGED ONLY")
        print("Open / Read = IGNORED")
        print("=" * 70)

        while True:
            try:
                check_category_and_flag_changes(
                    namespace,
                    tracked,
                )

                check_new_sent_items(
                    namespace,
                    tracked,
                    seen_sent_ids,
                )

            except Exception as error:
                print(f"⚠️ Tracker error: {error}")

            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        print("\n🛑 Phase 3 stopped.")

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
