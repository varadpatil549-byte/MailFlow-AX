import json
import time
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from ollama import chat

# Desktop front end
import io
import contextlib
import threading
from PySide6.QtCore import QObject, QThread, Signal, Slot, Qt, QTimer
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QGridLayout,
    QTabWidget,
    QScrollArea,
    QProgressBar,
    QSizePolicy,
    QWidget,
)


# ============================================================
# MAILFLOW AX
# FINAL SINGLE-PIPELINE APPLICATION
#
# Outlook -> Local AI -> Category -> Allocation -> Move
#          -> Single Excel Log -> Reply/Forward Tracking
#
# IMPORTANT RECOVERY BEHAVIOR
# ---------------------------
# First run establishes a deployment baseline and ignores historical mail.
# Later starts reconcile only emails received since the last checkpoint.
# This lets AX recover messages that arrived while AX was offline.
# Message identification uses PR_SEARCH_KEY when available, then Internet
# Message ID, with EntryID only as a last-resort fallback.
#
# Actioned = YES ONLY when Reply / Reply All / Forward is sent.
# Category and flag changes are logged but do not make Actioned.
# ============================================================


# ============================================================
# SETTINGS
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_FOLDER = PROJECT_ROOT / "data"

EXCEL_FILE = DATA_FOLDER / "MailFlow_AX_Final.xlsx"

# Leave blank to use the default Outlook mailbox.
# For a shared mailbox later, enter its exact Outlook store name.
MAILBOX_NAME = ""

INBOX_NAME = "Inbox"
SENT_ITEMS_NAME = "Sent Items"

MAIL_ITEM_CLASS = 43

POLL_INTERVAL = 2
MAX_RUNTIME_SCAN = 100
MAX_STARTUP_SCAN = None
MAX_BODY_CHARS = 5000

AI_MODEL = "qwen3:4b-instruct"

DEFAULT_USERS = [
    "User1",
    "User2",
    "User3",
]

CATEGORIES = [
    "Invoice",
    "Payment",
    "Vendor Query",
    "PO / Procurement",
    "GRIR",
    "Access / System",
    "Process Issue",
    "General",
    "Other",
]

# PR_SEARCH_KEY / PidTagSearchKey
# Stable across Outlook folder moves.
SEARCH_KEY_SCHEMA = (
    "http://schemas.microsoft.com/mapi/proptag/0x300B0102"
)

# Fallback stable identifier: Internet Message ID.
INTERNET_MESSAGE_ID_SCHEMAS = [
    "http://schemas.microsoft.com/mapi/proptag/0x1035001F",
    "http://schemas.microsoft.com/mapi/proptag/0x1035001E",
]

# Runtime checkpoint lives inside Client_Config.
STATE_BASELINE = "Baseline_Initialized"
STATE_LAST_CHECK = "Last_Check_Time"


# ============================================================
# EXCEL SCHEMA
# ============================================================

LOG_HEADERS = [
    "Ticket_ID",
    "Message_ID",
    "Received_Time",
    "Sender_Name",
    "Sender_Email",
    "Subject",
    "Body_Preview",
    "AI_Category",
    "AI_Confidence",
    "Assigned_User",
    "Destination_Folder",
    "Allocation_Time",
    "Allocation_Status",
    "Category_Changed",
    "Category_Change_Time",
    "Current_Categories",
    "Flag_Changed",
    "Flag_Change_Time",
    "Current_Flag",
    "Actioned",
    "Final_Action_Type",
    "Final_Action_Time",
    "Reply_Forward_Subject",
    "Overall_Status",
    "Last_Checked",
    "Notes",
]

CONFIG_HEADERS = [
    "Category",
    "Allocation_Method",
    "User1",
    "User2",
    "User3",
]


# ============================================================
# GENERAL HELPERS
# ============================================================

def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_datetime(value):
    """
    Convert Outlook/Excel datetimes into naive local datetimes.
    This prevents comparisons between offset-aware and
    offset-naive datetimes.
    """
    if value is None:
        return None

    if isinstance(value, datetime):
        dt = value
    else:
        try:
            dt = datetime.fromisoformat(str(value))
        except Exception:
            return None

    if dt.tzinfo is not None:
        try:
            dt = dt.astimezone().replace(tzinfo=None)
        except Exception:
            dt = dt.replace(tzinfo=None)

    return dt


def normalize_subject(subject):
    value = str(subject or "").strip()

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


def sent_action_from_subject(subject):
    value = str(subject or "").strip().lower()

    if value.startswith("re:"):
        return "Reply / Reply All"

    if value.startswith("fw:") or value.startswith("fwd:"):
        return "Forward"

    return None


def normalize_search_key(value):
    """
    Convert Outlook's PR_SEARCH_KEY result to a stable string.
    """
    if value is None:
        return ""

    if isinstance(value, bytes):
        return value.hex().upper()

    if isinstance(value, bytearray):
        return bytes(value).hex().upper()

    if isinstance(value, memoryview):
        return bytes(value).hex().upper()

    # Some COM environments may expose a tuple/list of integers.
    if isinstance(value, (tuple, list)):
        try:
            return bytes(value).hex().upper()
        except Exception:
            pass

    return str(value)


# ============================================================
# STABLE MESSAGE IDENTIFIER
# ============================================================

def get_message_id(item):
    """
    Return a move-stable identifier.

    Priority: PR_SEARCH_KEY -> Internet Message ID -> EntryID fallback.
    """
    try:
        raw_key = item.PropertyAccessor.GetProperty(
            SEARCH_KEY_SCHEMA
        )
        key = normalize_search_key(raw_key)
        if key:
            return f"SEARCH:{key}"
    except Exception:
        pass

    for schema in INTERNET_MESSAGE_ID_SCHEMAS:
        try:
            value = str(
                item.PropertyAccessor.GetProperty(schema) or ""
            ).strip()
            if value:
                return f"IMID:{value}"
        except Exception:
            continue

    entry_id = getattr(item, "EntryID", None)
    if entry_id:
        print(
            "⚠️ Stable MAPI IDs unavailable for one email; "
            "using EntryID as last-resort identifier."
        )
        return f"ENTRY:{entry_id}"

    raise RuntimeError(
        "No usable Outlook message identifier was available."
    )


# ============================================================
# EXCEL INITIALIZATION / VALIDATION
# ============================================================

def create_final_workbook():
    DATA_FOLDER.mkdir(exist_ok=True)

    if EXCEL_FILE.exists():
        return

    workbook = Workbook()

    log_sheet = workbook.active
    log_sheet.title = "MailFlow_Log"
    log_sheet.append(LOG_HEADERS)

    config_sheet = workbook.create_sheet("Client_Config")
    config_sheet.append(CONFIG_HEADERS)

    for category in CATEGORIES:
        config_sheet.append([
            category,
            "Round Robin",
            "User1",
            "User2",
            "User3",
        ])

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
    )

    for sheet in (log_sheet, config_sheet):

        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        sheet.freeze_panes = "A2"

    widths = {
        "A": 14,
        "B": 44,
        "C": 22,
        "D": 24,
        "E": 34,
        "F": 45,
        "G": 70,
        "H": 22,
        "I": 16,
        "J": 18,
        "K": 22,
        "L": 22,
        "M": 20,
        "N": 16,
        "O": 22,
        "P": 22,
        "Q": 14,
        "R": 22,
        "S": 15,
        "T": 12,
        "U": 24,
        "V": 22,
        "W": 45,
        "X": 18,
        "Y": 22,
        "Z": 45,
    }

    for column, width in widths.items():
        log_sheet.column_dimensions[column].width = width

    config_sheet.column_dimensions["A"].width = 24
    config_sheet.column_dimensions["B"].width = 20
    config_sheet.column_dimensions["C"].width = 18
    config_sheet.column_dimensions["D"].width = 18
    config_sheet.column_dimensions["E"].width = 18

    for row in log_sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    log_table = Table(
        displayName="MailFlowLogTable",
        ref=f"A1:Z{max(2, log_sheet.max_row)}",
    )

    log_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    log_sheet.add_table(log_table)

    config_table = Table(
        displayName="ClientConfigTable",
        ref=f"A1:E{config_sheet.max_row}",
    )

    config_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    config_sheet.add_table(config_table)

    workbook.save(EXCEL_FILE)
    workbook.close()

    print(f"✅ Final workbook created: {EXCEL_FILE}")


def validate_final_workbook():
    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        if "MailFlow_Log" not in workbook.sheetnames:
            raise ValueError(
                "MailFlow_AX_Final.xlsx is missing MailFlow_Log."
            )

        if "Client_Config" not in workbook.sheetnames:
            raise ValueError(
                "MailFlow_AX_Final.xlsx is missing Client_Config."
            )

        headers = [
            workbook["MailFlow_Log"].cell(
                1,
                column,
            ).value
            for column in range(
                1,
                len(LOG_HEADERS) + 1,
            )
        ]

        if headers != LOG_HEADERS:
            raise ValueError(
                "The existing MailFlow_Log headers do not "
                "match the final MailFlow AX schema. "
                "Use a fresh MailFlow_AX_Final.xlsx."
            )

    finally:
        workbook.close()


# ============================================================
# CLIENT CONFIGURATION
# ============================================================

def load_client_config():
    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        sheet = workbook["Client_Config"]

        rules = {}

        for row in sheet.iter_rows(
            min_row=2,
            values_only=True,
        ):
            category = str(
                row[0] or ""
            ).strip()

            if not category:
                continue

            method = str(
                row[1] or "Round Robin"
            ).strip()

            users = [
                str(value).strip()
                for value in row[2:5]
                if value
            ]

            rules[category] = {
                "method": method,
                "users": users or DEFAULT_USERS,
            }

        return rules

    finally:
        workbook.close()


def ensure_runtime_state_area():
    """Create runtime checkpoint cells on Client_Config."""
    workbook = load_workbook(EXCEL_FILE)
    try:
        sheet = workbook["Client_Config"]
        sheet["G1"] = "Runtime_State"
        sheet["H1"] = "Value"
        sheet["G2"] = STATE_BASELINE
        sheet["G3"] = STATE_LAST_CHECK
        if sheet["H2"].value is None:
            sheet["H2"] = "NO"
        if sheet["H3"].value is None:
            sheet["H3"] = ""
        for cell in (sheet["G1"], sheet["H1"]):
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        sheet.column_dimensions["G"].width = 24
        sheet.column_dimensions["H"].width = 28
        workbook.save(EXCEL_FILE)
    finally:
        workbook.close()


def get_runtime_state():
    workbook = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    try:
        sheet = workbook["Client_Config"]
        return {
            str(sheet.cell(r, 7).value).strip(): sheet.cell(r, 8).value
            for r in range(2, 20)
            if sheet.cell(r, 7).value
        }
    finally:
        workbook.close()


def set_runtime_state(**values):
    workbook = load_workbook(EXCEL_FILE)
    try:
        sheet = workbook["Client_Config"]
        key_rows = {}
        for r in range(2, 20):
            key = str(sheet.cell(r, 7).value or "").strip()
            if key:
                key_rows[key] = r
        for key, value in values.items():
            row = key_rows.get(key)
            if row is None:
                row = max(4, sheet.max_row + 1)
                sheet.cell(row, 7).value = key
            sheet.cell(row, 8).value = value
        workbook.save(EXCEL_FILE)
    finally:
        workbook.close()


def parse_checkpoint(value):
    return normalize_datetime(value)


def get_max_logged_received_time():
    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        sheet = workbook["MailFlow_Log"]
        latest = None

        for row in range(
            2,
            sheet.max_row + 1,
        ):
            if str(
                sheet.cell(row, 13).value or ""
            ) != "Moved":
                continue

            raw = sheet.cell(
                row,
                3,
            ).value

            if not raw:
                continue

            value = normalize_datetime(
                raw
            )

            if value is None:
                continue

            if latest is None or value > latest:
                latest = value

        return latest

    finally:
        workbook.close()


# ============================================================
# OUTLOOK
# ============================================================

def connect_outlook():
    try:
        outlook = win32com.client.GetActiveObject(
            "Outlook.Application"
        )
    except Exception:
        outlook = win32com.client.Dispatch(
            "Outlook.Application"
        )

    print("✅ Connected to Classic Outlook")

    return outlook.GetNamespace("MAPI")


def get_mailbox_root(namespace):
    if not MAILBOX_NAME.strip():
        return None

    try:
        return namespace.Folders[
            MAILBOX_NAME
        ]
    except Exception as error:
        raise RuntimeError(
            f"Mailbox '{MAILBOX_NAME}' was not found "
            "in the Outlook profile."
        ) from error


def get_mail_folders(namespace):
    root = get_mailbox_root(namespace)

    if root is None:
        inbox = namespace.GetDefaultFolder(6)
        sent = namespace.GetDefaultFolder(5)
        return inbox, sent

    try:
        inbox = root.Folders[INBOX_NAME]
    except Exception as error:
        raise RuntimeError(
            f"Inbox was not found under mailbox "
            f"'{MAILBOX_NAME}'."
        ) from error

    try:
        sent = root.Folders[SENT_ITEMS_NAME]
    except Exception as error:
        raise RuntimeError(
            f"'{SENT_ITEMS_NAME}' was not found under "
            f"mailbox '{MAILBOX_NAME}'."
        ) from error

    return inbox, sent


def is_mail(item):
    return getattr(
        item,
        "Class",
        None,
    ) == MAIL_ITEM_CLASS


def extract_email(item):
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
        "message_id": get_message_id(item),
        "subject": item.Subject or "(No subject)",
        "sender_name": (
            item.SenderName or "(Unknown sender)"
        ),
        "sender_email": sender_email,
        "received_time": item.ReceivedTime,
        "body": body[:MAX_BODY_CHARS],
        "conversation_id": (
            getattr(
                item,
                "ConversationID",
                "",
            )
            or ""
        ),
        "conversation_topic": (
            getattr(
                item,
                "ConversationTopic",
                "",
            )
            or ""
        ),
        "categories": item.Categories or "",
        "flag_status": getattr(
            item,
            "FlagStatus",
            None,
        ),
    }


# ============================================================
# LOCAL AI
# ============================================================

def classify_email(subject, body):
    prompt = f"""
You are an enterprise email classification engine.

Classify this business email into EXACTLY ONE category:

{", ".join(CATEGORIES)}

Rules:
- Return ONLY valid JSON.
- category must be exactly one of the allowed categories.
- confidence must be an integer from 0 to 100.
- Do not invent a category.
- Do not return a reason.

SUBJECT:
{subject}

BODY:
{body}
"""

    response = chat(
        model=AI_MODEL,
        messages=[
            {
                "role": "user",
                "content": prompt,
            }
        ],
        format={
            "type": "object",
            "properties": {
                "category": {
                    "type": "string",
                    "enum": CATEGORIES,
                },
                "confidence": {
                    "type": "integer",
                    "minimum": 0,
                    "maximum": 100,
                },
            },
            "required": [
                "category",
                "confidence",
            ],
        },
        options={
            "temperature": 0,
        },
    )

    result = json.loads(
        response.message.content
    )

    category = result.get("category")
    confidence = int(
        result.get("confidence", 0)
    )

    if category not in CATEGORIES:
        raise ValueError(
            f"Invalid AI category: {category}"
        )

    return {
        "category": category,
        "confidence": confidence,
    }


# ============================================================
# LOG HELPERS
# ============================================================

def find_log_row(
    sheet,
    message_id,
):
    for row in range(
        2,
        sheet.max_row + 1,
    ):
        if (
            str(
                sheet.cell(row, 2).value or ""
            )
            == str(message_id)
        ):
            return row

    return None


def get_next_ticket(
    sheet,
):
    highest = 0

    for row in range(
        2,
        sheet.max_row + 1,
    ):
        value = str(
            sheet.cell(row, 1).value or ""
        )

        if value.startswith("MFAX-"):
            try:
                highest = max(
                    highest,
                    int(
                        value.split("-")[1]
                    ),
                )
            except Exception:
                pass

    return f"MFAX-{highest + 1:05d}"


def load_processed_message_ids():
    """
    Return the Message_ID values for successfully moved
    messages.

    Error rows are deliberately NOT treated as processed,
    so a failed message remains eligible for retry.
    """

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        sheet = workbook[
            "MailFlow_Log"
        ]

        processed = set()

        for row in range(
            2,
            sheet.max_row + 1,
        ):
            message_id = str(
                sheet.cell(row, 2).value or ""
            )

            allocation_status = str(
                sheet.cell(row, 13).value or ""
            )

            if (
                message_id
                and allocation_status == "Moved"
            ):
                processed.add(
                    message_id
                )

        return processed

    finally:
        workbook.close()


def upsert_log_record(
    email,
    ai_result,
    assigned_user,
    destination_folder,
    allocation_status,
    overall_status,
    notes="",
):
    """
    Insert a new log row or update the existing row for the
    same stable Message_ID.

    This prevents duplicate Excel rows when an email is retried.
    """

    workbook = load_workbook(
        EXCEL_FILE
    )

    try:
        sheet = workbook[
            "MailFlow_Log"
        ]

        row = find_log_row(
            sheet,
            email["message_id"],
        )

        if row is None:
            ticket_id = get_next_ticket(
                sheet
            )

            sheet.append([
                ticket_id,
                email["message_id"],
                str(email["received_time"]),
                email["sender_name"],
                email["sender_email"],
                email["subject"],
                email["body"],
                ai_result.get(
                    "category",
                    "",
                ),
                ai_result.get(
                    "confidence",
                    "",
                ),
                assigned_user,
                destination_folder,
                now_text(),
                allocation_status,
                "NO",
                "",
                email.get(
                    "categories",
                    "",
                ),
                "NO",
                "",
                str(
                    email.get(
                        "flag_status",
                        "",
                    )
                ),
                "NO",
                "",
                "",
                "",
                overall_status,
                now_text(),
                notes,
            ])

            workbook.save(
                EXCEL_FILE
            )

            return ticket_id

        # Existing row: update processing/allocation fields.
        sheet.cell(
            row,
            8,
        ).value = ai_result.get(
            "category",
            "",
        )

        sheet.cell(
            row,
            9,
        ).value = ai_result.get(
            "confidence",
            "",
        )

        sheet.cell(
            row,
            10,
        ).value = assigned_user

        sheet.cell(
            row,
            11,
        ).value = destination_folder

        sheet.cell(
            row,
            12,
        ).value = now_text()

        sheet.cell(
            row,
            13,
        ).value = allocation_status

        sheet.cell(
            row,
            24,
        ).value = overall_status

        sheet.cell(
            row,
            25,
        ).value = now_text()

        sheet.cell(
            row,
            26,
        ).value = notes

        workbook.save(
            EXCEL_FILE
        )

        return sheet.cell(
            row,
            1,
        ).value

    finally:
        workbook.close()


# ============================================================
# ROUND-ROBIN ALLOCATION
# ============================================================

def count_previous_allocations(
    sheet,
    eligible_users,
):
    """
    Count successful allocations across the eligible-user
    pool, regardless of AI category.

    This prevents:
        Invoice -> User1
        Payment -> User1
        PO -> User1

    and instead creates a shared sequence:
        User1 -> User2 -> User3 -> User1 ...
    """

    count = 0

    for row in range(
        2,
        sheet.max_row + 1,
    ):
        allocation_status = str(
            sheet.cell(
                row,
                13,
            ).value or ""
        )

        assigned_user = str(
            sheet.cell(
                row,
                10,
            ).value or ""
        )

        if (
            allocation_status == "Moved"
            and assigned_user in eligible_users
        ):
            count += 1

    return count


def choose_user(
    category,
    config,
):
    rule = config.get(
        category,
        {
            "method": "Round Robin",
            "users": DEFAULT_USERS,
        },
    )

    users = [
        str(user).strip()
        for user in rule.get(
            "users",
            DEFAULT_USERS,
        )
        if user
    ]

    if not users:
        raise ValueError(
            f"No eligible users configured for "
            f"category '{category}'."
        )

    method = str(
        rule.get(
            "method",
            "Round Robin",
        )
    ).strip().lower()

    if (
        method == "direct"
        or len(users) == 1
    ):
        return users[0]

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        sheet = workbook[
            "MailFlow_Log"
        ]

        previous = count_previous_allocations(
            sheet,
            users,
        )

        return users[
            previous % len(users)
        ]

    finally:
        workbook.close()


# ============================================================
# INBOX RECONCILIATION
# ============================================================

def get_inbox_candidates(
    inbox,
    processed_ids,
    since_time=None,
    max_scan=None,
):
    """
    Scan newest -> oldest and return only unprocessed emails
    at/after the last checkpoint.

    All datetime values are normalized before comparison so
    Outlook timezone-aware ReceivedTime values can be safely
    compared with Python/Excel checkpoint values.
    """
    since_time = normalize_datetime(
        since_time
    )

    items = inbox.Items
    items.Sort(
        "[ReceivedTime]",
        True,
    )

    candidates = []
    scanned = 0

    for item in items:

        if (
            max_scan is not None
            and scanned >= max_scan
        ):
            break

        if not is_mail(item):
            continue

        scanned += 1

        try:
            received_time = normalize_datetime(
                item.ReceivedTime
            )

            if received_time is None:
                print(
                    "⚠️ Skipping Inbox email with "
                    "unreadable ReceivedTime."
                )
                continue

            if (
                since_time is not None
                and received_time < since_time
            ):
                break

            message_id = get_message_id(
                item
            )

        except Exception as error:
            print(
                f"⚠️ Could not identify Inbox email: "
                f"{error}"
            )
            continue

        if message_id not in processed_ids:
            candidates.append(item)

    candidates.reverse()
    return candidates


# ============================================================
# NEW EMAIL -> AI -> ALLOCATE -> MOVE
# ============================================================

def process_new_email(
    item,
    inbox,
    config,
):
    email = extract_email(
        item
    )

    print("\n" + "=" * 75)
    print(
        "📨 NEW / RECOVERED EMAIL"
    )
    print("=" * 75)

    print(
        f"Subject : {email['subject']}"
    )
    print(
        f"Sender  : {email['sender_email']}"
    )
    print(
        f"Received: {email['received_time']}"
    )

    # --------------------------------------------------------
    # AI
    # --------------------------------------------------------

    print(
        "\n🤖 AI classifying..."
    )

    try:
        ai_result = classify_email(
            email["subject"],
            email["body"],
        )

        print(
            f"✅ Category   : "
            f"{ai_result['category']}"
        )

        print(
            f"✅ Confidence : "
            f"{ai_result['confidence']}%"
        )

    except Exception as error:

        ticket_id = upsert_log_record(
            email=email,
            ai_result={
                "category": "",
                "confidence": "",
            },
            assigned_user="",
            destination_folder="",
            allocation_status="AI Error",
            overall_status="Pending",
            notes=str(error),
        )

        print(
            f"❌ AI classification failed: "
            f"{error}"
        )

        print(
            f"⚠️ Logged for retry: {ticket_id}"
        )

        return False

    # --------------------------------------------------------
    # ALLOCATION
    # --------------------------------------------------------

    try:
        assigned_user = choose_user(
            ai_result["category"],
            config,
        )

        destination = inbox.Folders[
            assigned_user
        ]

        print(
            f"➡️ Assigned: "
            f"{assigned_user}"
        )

        print(
            f"📁 Destination: "
            f"{destination.Name}"
        )

    except Exception as error:

        ticket_id = upsert_log_record(
            email=email,
            ai_result=ai_result,
            assigned_user="",
            destination_folder="",
            allocation_status="Allocation Error",
            overall_status="Pending",
            notes=str(error),
        )

        print(
            f"❌ Allocation failed: "
            f"{error}"
        )

        print(
            f"⚠️ Logged for retry: {ticket_id}"
        )

        return False

    # --------------------------------------------------------
    # MOVE
    # --------------------------------------------------------

    try:
        # Log BEFORE move so a restart cannot forget the
        # message if the process stops between steps.
        ticket_id = upsert_log_record(
            email=email,
            ai_result=ai_result,
            assigned_user=assigned_user,
            destination_folder=destination.Name,
            allocation_status="Pending Move",
            overall_status="Processing",
        )

        print(
            f"🎫 Ticket: {ticket_id}"
        )

        item.Move(
            destination
        )

        # The stable Message_ID does not need to change.
        # EntryID is intentionally not used for tracking.

        upsert_log_record(
            email=email,
            ai_result=ai_result,
            assigned_user=assigned_user,
            destination_folder=destination.Name,
            allocation_status="Moved",
            overall_status="Allocated",
        )

        print(
            "✅ EMAIL MOVED SUCCESSFULLY"
        )

        return True

    except Exception as error:

        upsert_log_record(
            email=email,
            ai_result=ai_result,
            assigned_user=assigned_user,
            destination_folder=destination.Name,
            allocation_status="Move Error",
            overall_status="Pending",
            notes=str(error),
        )

        print(
            f"❌ Move failed: {error}"
        )

        return False


# ============================================================
# ACTION TRACKING
# ============================================================

def load_open_action_records():
    """
    Load successfully moved but not-yet-actioned emails.
    """

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        sheet = workbook[
            "MailFlow_Log"
        ]

        records = {}

        for row in range(
            2,
            sheet.max_row + 1,
        ):

            message_id = str(
                sheet.cell(
                    row,
                    2,
                ).value or ""
            )

            assigned_user = str(
                sheet.cell(
                    row,
                    10,
                ).value or ""
            )

            allocation_status = str(
                sheet.cell(
                    row,
                    13,
                ).value or ""
            )

            actioned = str(
                sheet.cell(
                    row,
                    20,
                ).value or "NO"
            )

            if (
                message_id
                and assigned_user
                and allocation_status == "Moved"
                and actioned != "YES"
            ):

                subject = str(
                    sheet.cell(
                        row,
                        6,
                    ).value or ""
                )

                records[
                    message_id
                ] = {
                    "message_id": message_id,
                    "assigned_user": assigned_user,
                    "subject": subject,
                    "normalized_subject": (
                        normalize_subject(
                            subject
                        )
                    ),
                    "received_time": (
                        sheet.cell(
                            row,
                            3,
                        ).value
                    ),
                    "conversation_id": "",
                    "conversation_topic": "",
                    "categories": str(
                        sheet.cell(
                            row,
                            16,
                        ).value or ""
                    ),
                    "flag": sheet.cell(
                        row,
                        19,
                    ).value,
                }

        return records

    finally:
        workbook.close()


def refresh_action_records(
    inbox,
    records,
):
    """
    Locate tracked emails inside their assigned folders
    using the stable MAPI Search Key.
    """

    for folder_name in DEFAULT_USERS:

        try:
            folder = inbox.Folders[
                folder_name
            ]
        except Exception:
            continue

        for item in folder.Items:

            if not is_mail(item):
                continue

            try:
                message_id = get_message_id(
                    item
                )
            except Exception:
                continue

            if message_id not in records:
                continue

            record = records[
                message_id
            ]

            record[
                "conversation_id"
            ] = (
                getattr(
                    item,
                    "ConversationID",
                    "",
                )
                or ""
            )

            record[
                "conversation_topic"
            ] = (
                getattr(
                    item,
                    "ConversationTopic",
                    "",
                )
                or ""
            )

            record[
                "categories"
            ] = (
                item.Categories or ""
            )

            record[
                "flag"
            ] = getattr(
                item,
                "FlagStatus",
                None,
            )


def update_log(
    message_id,
    updates,
):
    workbook = load_workbook(
        EXCEL_FILE
    )

    try:
        sheet = workbook[
            "MailFlow_Log"
        ]

        row = find_log_row(
            sheet,
            message_id,
        )

        if row is None:
            return

        for column, value in updates.items():
            sheet.cell(
                row,
                column,
            ).value = value

        sheet.cell(
            row,
            25,
        ).value = now_text()

        workbook.save(
            EXCEL_FILE
        )

    finally:
        workbook.close()


def check_category_and_flag_changes(
    inbox,
    records,
):
    """
    Category/flag changes are tracked.
    They do NOT make Actioned = YES.
    """

    for message_id, record in list(
        records.items()
    ):

        try:
            folder = inbox.Folders[
                record["assigned_user"]
            ]
        except Exception:
            continue

        found = None

        for item in folder.Items:

            if not is_mail(item):
                continue

            try:
                if (
                    get_message_id(item)
                    == message_id
                ):
                    found = item
                    break
            except Exception:
                continue

        if found is None:
            continue

        current_categories = (
            found.Categories or ""
        )

        current_flag = getattr(
            found,
            "FlagStatus",
            None,
        )

        updates = {}

        if (
            current_categories
            != record.get(
                "categories",
                "",
            )
        ):

            record[
                "categories"
            ] = current_categories

            updates[14] = "YES"
            updates[15] = now_text()
            updates[16] = current_categories

            print(
                f"🏷️ Category changed: "
                f"{record['subject']}"
            )

        if (
            current_flag
            != record.get(
                "flag",
                None,
            )
        ):

            record[
                "flag"
            ] = current_flag

            updates[17] = "YES"
            updates[18] = now_text()
            updates[19] = str(
                current_flag
            )

            print(
                f"🚩 Flag changed: "
                f"{record['subject']}"
            )

        if updates:
            update_log(
                message_id,
                updates,
            )


def find_action_match(
    sent_item,
    records,
):
    sent_subject = (
        sent_item.Subject or ""
    ).strip()

    action_type = sent_action_from_subject(
        sent_subject
    )

    if not action_type:
        return None, None

    normalized_sent_subject = normalize_subject(
        sent_subject
    )

    sent_conversation_id = (
        getattr(
            sent_item,
            "ConversationID",
            "",
        )
        or ""
    )

    sent_conversation_topic = (
        getattr(
            sent_item,
            "ConversationTopic",
            "",
        )
        or ""
    )

    for record in records.values():

        if (
            sent_conversation_id
            and record.get(
                "conversation_id"
            )
            and (
                sent_conversation_id
                == record[
                    "conversation_id"
                ]
            )
        ):
            return record, action_type

    for record in records.values():

        if (
            sent_conversation_topic
            and record.get(
                "conversation_topic"
            )
            and (
                sent_conversation_topic.lower()
                == record[
                    "conversation_topic"
                ].lower()
            )
        ):
            return record, action_type

    matches = [
        record
        for record in records.values()
        if (
            record[
                "normalized_subject"
            ]
            == normalized_sent_subject
        )
    ]

    if len(matches) == 1:
        return (
            matches[0],
            action_type,
        )

    return None, None


def build_seen_sent_ids(
    sent_folder,
):
    seen = set()

    items = sent_folder.Items

    items.Sort(
        "[SentOn]",
        True,
    )

    count = 0

    for item in items:

        if count >= MAX_RUNTIME_SCAN:
            break

        if not is_mail(item):
            continue

        count += 1

        try:
            if item.EntryID:
                seen.add(
                    item.EntryID
                )
        except Exception:
            pass

    return seen


def check_new_sent_items(
    sent_folder,
    records,
    seen_sent_ids,
):
    """
    Reply / Reply All / Forward are the only events that
    set Actioned = YES.
    """

    items = sent_folder.Items

    items.Sort(
        "[SentOn]",
        True,
    )

    count = 0

    for item in items:

        if count >= MAX_RUNTIME_SCAN:
            break

        if not is_mail(item):
            continue

        count += 1

        try:
            sent_id = item.EntryID

            if not sent_id:
                continue

            if sent_id in seen_sent_ids:
                continue

            seen_sent_ids.add(
                sent_id
            )

            record, action_type = (
                find_action_match(
                    item,
                    records,
                )
            )

            if record is None:
                continue

            message_id = record[
                "message_id"
            ]

            update_log(
                message_id,
                {
                    20: "YES",
                    21: action_type,
                    22: now_text(),
                    23: item.Subject or "",
                    24: "Actioned",
                },
            )

            print("\n" + "=" * 75)
            print(
                "✅ USER ACTION DETECTED"
            )
            print("=" * 75)
            print(
                f"Original : "
                f"{record['subject']}"
            )
            print(
                f"User     : "
                f"{record['assigned_user']}"
            )
            print(
                f"Action   : "
                f"{action_type}"
            )
            print(
                "Actioned : YES"
            )
            print("=" * 75)

            records.pop(
                message_id,
                None,
            )

        except Exception:
            continue



# ============================================================
# PIPELINE ENGINE
# ============================================================

def run_pipeline(stop_event):
    print("\n" + "=" * 75)
    print("MAILFLOW AX - FINAL SINGLE PIPELINE")
    print("OUTLOOK -> AI -> ALLOCATION -> MOVE -> ACTION")
    print("=" * 75)

    create_final_workbook()
    validate_final_workbook()
    ensure_runtime_state_area()

    config = load_client_config()
    pythoncom.CoInitialize()

    try:
        namespace = connect_outlook()
        inbox, sent_folder = get_mail_folders(namespace)

        print(f"✅ Monitoring Inbox: {inbox.Name}")
        print(f"✅ Monitoring Sent Items: {sent_folder.Name}")
        print("✅ Date/time normalization enabled")

        # ----------------------------------------------------
        # FIRST RUN / RECOVERY CHECKPOINT
        # ----------------------------------------------------
        state = get_runtime_state()
        baseline_initialized = str(
            state.get(STATE_BASELINE, "NO") or "NO"
        ).upper() == "YES"
        last_check = parse_checkpoint(
            state.get(STATE_LAST_CHECK)
        )

        if not baseline_initialized:
            baseline_time = datetime.now()
            logged_time = get_max_logged_received_time()
            if logged_time and logged_time > baseline_time:
                baseline_time = logged_time

            set_runtime_state(
                **{
                    STATE_BASELINE: "YES",
                    STATE_LAST_CHECK: baseline_time.isoformat(sep=" "),
                }
            )
            last_check = baseline_time

            print("🟢 First-run baseline initialized.")
            print(f"   Baseline: {baseline_time}")
            print("   Existing historical Inbox mail will NOT be processed.")

        processed_ids = load_processed_message_ids()

        if last_check is None:
            last_check = get_max_logged_received_time()

        print(
            f"✅ Processed Message_IDs loaded from Excel: {len(processed_ids)}"
        )
        print(
            f"🔎 Recovery checkpoint: {last_check}"
        )

        recovery_candidates = get_inbox_candidates(
            inbox,
            processed_ids,
            since_time=last_check,
            max_scan=None,
        )

        print(
            f"🔎 Recovery scan found {len(recovery_candidates)} "
            "email(s) since the last checkpoint."
        )

        for item in recovery_candidates:
            if stop_event.is_set():
                break

            success = process_new_email(
                item,
                inbox,
                config,
            )
            if success:
                try:
                    processed_ids.add(get_message_id(item))
                except Exception:
                    pass

        if stop_event.is_set():
            print("\n🛑 MailFlow AX stopped.")
            return

        action_records = load_open_action_records()
        refresh_action_records(inbox, action_records)
        seen_sent_ids = build_seen_sent_ids(sent_folder)

        # Checkpoint AFTER recovery processing, not before.
        set_runtime_state(
            **{
                STATE_LAST_CHECK: datetime.now().isoformat(sep=" ")
            }
        )

        print(f"✅ Open action records: {len(action_records)}")
        print("\n🟢 MAILFLOW AX IS RUNNING")
        print("New Email -> AI -> Allocation -> Move -> Action")
        print("First run -> historical Inbox ignored")
        print("Restart -> emails since last checkpoint recovered")
        print("Actioned = Reply / Reply All / Forward ONLY")
        print("Open/Read = ignored | Category/Flag = logged only")
        print("\nPress STOP MONITORING to stop.")
        print("=" * 75)

        # ----------------------------------------------------
        # NORMAL RUNTIME
        # ----------------------------------------------------
        while not stop_event.is_set():
            cycle_start = datetime.now()
            config = load_client_config()
            processed_ids.update(load_processed_message_ids())

            runtime_candidates = get_inbox_candidates(
                inbox,
                processed_ids,
                since_time=last_check,
                max_scan=None,
            )

            for item in runtime_candidates:
                if stop_event.is_set():
                    break

                try:
                    message_id = get_message_id(item)
                except Exception:
                    message_id = None

                success = process_new_email(
                    item,
                    inbox,
                    config,
                )

                if success and message_id:
                    processed_ids.add(message_id)

            if stop_event.is_set():
                break

            action_records = load_open_action_records()
            refresh_action_records(inbox, action_records)
            check_category_and_flag_changes(inbox, action_records)
            check_new_sent_items(
                sent_folder,
                action_records,
                seen_sent_ids,
            )

            # Update checkpoint only after this cycle is done.
            last_check = cycle_start
            set_runtime_state(
                **{
                    STATE_LAST_CHECK: last_check.isoformat(sep=" ")
                }
            )

            stop_event.wait(POLL_INTERVAL)

        print("\n🛑 MailFlow AX stopped.")

    except Exception as error:
        print(f"\n❌ MailFlow AX engine error: {error}")
        raise
    finally:
        pythoncom.CoUninitialize()


# ============================================================
# DESKTOP FRONT END
# ============================================================

class SignalWriter(io.TextIOBase):
    """Captures normal print() output and sends it to the GUI."""

    def __init__(self, signal):
        super().__init__()
        self.signal = signal
        self._buffer = ""

    def write(self, value):
        if not value:
            return 0

        self._buffer += str(value)

        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            self.signal.emit(line.rstrip("\r"))

        return len(value)

    def flush(self):
        if self._buffer:
            self.signal.emit(self._buffer.rstrip("\r"))
            self._buffer = ""


class PipelineWorker(QObject):
    output = Signal(str)
    finished = Signal()
    failed = Signal(str)

    def __init__(self):
        super().__init__()
        self.stop_event = threading.Event()

    @Slot()
    def run(self):
        writer = SignalWriter(self.output)

        try:
            with contextlib.redirect_stdout(writer):
                run_pipeline(self.stop_event)
        except Exception as error:
            self.failed.emit(str(error))
        finally:
            writer.flush()
            self.finished.emit()

    def stop(self):
        self.stop_event.set()


class MailFlowAXWindow(QMainWindow):
    RED = "#D71920"
    RED_DARK = "#A50F15"
    BG = "#0B0D0F"
    PANEL = "#121519"
    PANEL_2 = "#171B20"
    BORDER = "#2A3037"
    TEXT = "#F2F4F7"
    MUTED = "#9AA3AD"
    GREEN = "#22C55E"
    CONSOLE = "#080A0C"

    def __init__(self):
        super().__init__()
        self.thread = None
        self.worker = None
        self.is_running = False
        self.dashboard_timer = QTimer(self)
        self.dashboard_timer.setInterval(3000)
        self.dashboard_timer.timeout.connect(self.refresh_dashboard)

        self.setWindowTitle("MailFlow AX")
        self.setMinimumSize(1080, 720)
        self.resize(1220, 820)
        self.build_ui()

    def build_ui(self):
        root = QWidget()
        root.setObjectName("Root")
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(30, 24, 30, 24)
        outer.setSpacing(15)

        # ========================================================
        # HEADER
        # ========================================================
        header = QHBoxLayout()
        header.setSpacing(22)

        logo = QLabel()
        logo.setFixedSize(270, 165)
        logo.setAlignment(Qt.AlignCenter)

        # MailFlow AX official dark/neon logo
        # Stored inside the project's .venv folder as requested.
        logo_path = Path(
            r"C:\Users\varad.patil\Desktop\Mailflow Ax\.venv\MailFlow AX Neon AI Mailbox.png"
        )

        # Fallbacks keep the UI usable if the machine/path changes later.
        if not logo_path.exists():
            local_candidates = [
                Path(__file__).resolve().parent / "MailFlow AX Neon AI Mailbox.png",
                Path(__file__).resolve().parent / "MailFlow_AX_Logo_transparent.png",
                Path(__file__).resolve().parent / "MailFlow_AX_Logo.png",
            ]
            logo_path = next(
                (p for p in local_candidates if p.exists()),
                None,
            )

        if logo_path:
            pixmap = QPixmap(str(logo_path))
            logo.setPixmap(
                pixmap.scaled(
                    265,
                    160,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation,
                )
            )
        else:
            logo.setText("MAILFLOW AX")
            logo.setStyleSheet(
                "font-size: 26px; font-weight: 900; color: #F2F4F7;"
            )

        header.addWidget(logo)

        title_box = QVBoxLayout()
        title_box.setSpacing(5)

        title = QLabel("MailFlow Ax")
        title.setObjectName("Title")

        subtitle = QLabel("AI POWERED MAILBOX EXCHANGE")
        subtitle.setObjectName("Subtitle")

        description = QLabel(
            "Local AI  •  Intelligent Allocation  •  Automated Mail Movement  •  Action Tracking"
        )
        description.setObjectName("Description")
        description.setWordWrap(True)

        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        title_box.addSpacing(8)
        title_box.addWidget(description)
        title_box.addStretch()

        header.addLayout(title_box, 1)

        status_card = QFrame()
        status_card.setObjectName("StatusCard")
        status_card.setMinimumWidth(220)

        status_layout = QVBoxLayout(status_card)
        status_layout.setContentsMargins(20, 16, 20, 16)
        status_layout.setSpacing(8)

        status_heading = QLabel("SYSTEM STATUS")
        status_heading.setObjectName("StatusHeading")

        status_row = QHBoxLayout()
        self.status_dot = QLabel("●")
        self.status_dot.setObjectName("StatusDot")

        self.status_label = QLabel("NOT RUNNING")
        self.status_label.setObjectName("StatusText")

        status_row.addWidget(self.status_dot)
        status_row.addWidget(self.status_label)
        status_row.addStretch()

        status_layout.addWidget(status_heading)
        status_layout.addLayout(status_row)

        header.addWidget(status_card)
        outer.addLayout(header)

        # ========================================================
        # CONTROL BAR
        # ========================================================
        controls = QFrame()
        controls.setObjectName("ControlBar")

        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(18, 13, 18, 13)
        controls_layout.setSpacing(10)

        self.start_button = QPushButton("▶  START MONITORING")
        self.start_button.setObjectName("StartButton")
        self.start_button.clicked.connect(self.start_monitoring)

        self.stop_button = QPushButton("■  STOP MONITORING")
        self.stop_button.setObjectName("StopButton")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_monitoring)

        controls_layout.addWidget(self.start_button)
        controls_layout.addWidget(self.stop_button)
        controls_layout.addStretch()

        engine_label = QLabel("ENGINE  •  LOCAL AI / OLLAMA")
        engine_label.setObjectName("SmallInfo")
        controls_layout.addWidget(engine_label)

        outer.addWidget(controls)

        # ========================================================
        # INFO CARDS
        # ========================================================
        cards = QHBoxLayout()
        cards.setSpacing(10)

        for heading, value in [
            ("MAILBOX", "Classic Outlook"),
            ("AI ENGINE", "Local / Offline"),
            ("PIPELINE", "AI → Allocate → Move → Track"),
            ("LOG", "MailFlow_AX_Final.xlsx"),
        ]:
            cards.addWidget(self.make_info_card(heading, value))

        outer.addLayout(cards)

        # ========================================================
        # DASHBOARD + LIVE ACTIVITY
        # ========================================================
        self.tabs = QTabWidget()
        self.tabs.setObjectName("MainTabs")

        self.dashboard_page = self.build_dashboard_page()
        self.activity_page = self.build_activity_page()

        self.tabs.addTab(self.dashboard_page, "  DASHBOARD  ")
        self.tabs.addTab(self.activity_page, "  LIVE ACTIVITY  ")
        outer.addWidget(self.tabs, 1)

        self.dashboard_timer.start()
        self.refresh_dashboard()

        footer = QLabel(
            "MAILFLOW AX  •  ENTERPRISE MAIL AUTOMATION  •  ACE"
        )
        footer.setObjectName("Footer")
        outer.addWidget(footer)

        self.setStyleSheet(self.styles())

    def build_activity_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(4, 8, 4, 4)
        layout.setSpacing(8)

        header = QHBoxLayout()
        live_title = QLabel("LIVE ACTIVITY")
        live_title.setObjectName("SectionTitle")
        header.addWidget(live_title)
        header.addStretch()
        self.activity_status = QLabel("Waiting to start")
        self.activity_status.setObjectName("ActivityStatus")
        header.addWidget(self.activity_status)
        layout.addLayout(header)

        self.console = QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.console.setPlaceholderText("MailFlow AX activity will appear here...")
        self.console.setObjectName("Console")
        layout.addWidget(self.console, 1)
        return page

    def build_dashboard_page(self):
        page = QWidget()
        page.setObjectName("DashboardPage")
        outer = QVBoxLayout(page)
        outer.setContentsMargins(4, 8, 4, 4)
        outer.setSpacing(12)

        top = QHBoxLayout()
        dash_title = QLabel("OPERATIONS DASHBOARD")
        dash_title.setObjectName("DashboardTitle")
        top.addWidget(dash_title)
        top.addStretch()
        self.dashboard_updated = QLabel("Waiting for data...")
        self.dashboard_updated.setObjectName("DashboardUpdated")
        top.addWidget(self.dashboard_updated)
        refresh_btn = QPushButton("↻  REFRESH")
        refresh_btn.setObjectName("RefreshButton")
        refresh_btn.clicked.connect(self.refresh_dashboard)
        top.addWidget(refresh_btn)
        outer.addLayout(top)

        cards = QGridLayout()
        cards.setHorizontalSpacing(10)
        cards.setVerticalSpacing(10)
        self.kpi_widgets = {}
        kpis = [
            ("Total Emails", "total"),
            ("Allocated", "allocated"),
            ("Actioned", "actioned"),
            ("Pending", "pending"),
            ("Action Rate", "action_rate"),
            ("AI Confidence", "ai_conf"),
        ]
        for i, (label, key) in enumerate(kpis):
            card = self.make_kpi_card(label, "—")
            self.kpi_widgets[key] = card.findChild(QLabel, "KPIValue")
            cards.addWidget(card, 0, i)
        outer.addLayout(cards)

        content = QGridLayout()
        content.setHorizontalSpacing(12)
        content.setVerticalSpacing(12)

        self.category_panel = self.make_bar_panel("EMAIL VOLUME BY CATEGORY")
        self.user_panel = self.make_bar_panel("WORKLOAD BY ASSIGNED USER")
        self.aging_panel = self.make_bar_panel("PENDING EMAIL AGING")
        self.routing_panel = self.make_bar_panel("ROUTING DISTRIBUTION")

        content.addWidget(self.category_panel, 0, 0)
        content.addWidget(self.user_panel, 0, 1)
        content.addWidget(self.aging_panel, 1, 0)
        content.addWidget(self.routing_panel, 1, 1)
        outer.addLayout(content, 1)
        return page

    def make_kpi_card(self, label, value):
        frame = QFrame()
        frame.setObjectName("KPICard")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(14, 11, 14, 11)
        layout.setSpacing(5)
        h = QLabel(label.upper())
        h.setObjectName("KPIHeading")
        v = QLabel(value)
        v.setObjectName("KPIValue")
        layout.addWidget(h)
        layout.addWidget(v)
        return frame

    def make_bar_panel(self, title):
        frame = QFrame()
        frame.setObjectName("DashboardPanel")
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(15, 13, 15, 13)
        layout.setSpacing(8)
        heading = QLabel(title)
        heading.setObjectName("PanelHeading")
        layout.addWidget(heading)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 4, 0, 0)
        body_layout.setSpacing(7)
        frame._body = body_layout
        layout.addWidget(body, 1)
        return frame

    def clear_panel(self, panel):
        while panel._body.count():
            item = panel._body.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def add_bar_row(self, layout, label, value, maximum, suffix=""):
        row = QWidget()
        r = QVBoxLayout(row)
        r.setContentsMargins(0, 0, 0, 0)
        r.setSpacing(3)
        head = QHBoxLayout()
        name = QLabel(str(label))
        name.setObjectName("BarLabel")
        num = QLabel(f"{value}{suffix}")
        num.setObjectName("BarValue")
        head.addWidget(name)
        head.addStretch()
        head.addWidget(num)
        r.addLayout(head)
        bar = QProgressBar()
        bar.setRange(0, max(1, int(maximum)))
        bar.setValue(max(0, int(value)))
        bar.setTextVisible(False)
        bar.setObjectName("DashboardBar")
        r.addWidget(bar)
        layout.addWidget(row)

    def _read_dashboard_rows(self):
        if not EXCEL_FILE.exists():
            return []
        wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
        try:
            if "MailFlow_Log" in wb.sheetnames:
                ws = wb["MailFlow_Log"]
            elif "MailFlowLogTable" in wb.sheetnames:
                ws = wb["MailFlowLogTable"]
            else:
                return []
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(x) if x is not None else "" for x in rows[0]]
            return [dict(zip(headers, r)) for r in rows[1:] if any(v is not None for v in r)]
        finally:
            wb.close()

    def _parse_dashboard_datetime(self, value):
        if isinstance(value, datetime):
            return value.replace(tzinfo=None)
        if value is None:
            return None
        text = str(value).strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
        try:
            return datetime.fromisoformat(text.replace("Z", "")).replace(tzinfo=None)
        except Exception:
            return None

    def refresh_dashboard(self):
        try:
            rows = self._read_dashboard_rows()
            total = len(rows)
            allocated = sum(1 for r in rows if str(r.get("Allocation_Status", "")).strip().lower() == "allocated")
            actioned = sum(1 for r in rows if str(r.get("Actioned", "")).strip().lower() == "yes")
            pending_rows = [r for r in rows if str(r.get("Actioned", "")).strip().lower() != "yes"]
            pending = len(pending_rows)

            def pct(a, b):
                return (a / b * 100) if b else 0

            confidences = []
            for r in rows:
                try:
                    c = float(r.get("AI_Confidence"))
                    confidences.append(c * 100 if c <= 1 else c)
                except (TypeError, ValueError):
                    pass
            ai_conf = sum(confidences) / len(confidences) if confidences else 0

            self.kpi_widgets["total"].setText(str(total))
            self.kpi_widgets["allocated"].setText(str(allocated))
            self.kpi_widgets["actioned"].setText(str(actioned))
            self.kpi_widgets["pending"].setText(str(pending))
            self.kpi_widgets["action_rate"].setText(f"{pct(actioned, total):.2f}%")
            self.kpi_widgets["ai_conf"].setText(f"{ai_conf:.2f}%")

            categories = {}
            for r in rows:
                key = str(r.get("AI_Category") or "Unclassified")
                categories[key] = categories.get(key, 0) + 1
            users = {}
            for r in rows:
                key = str(r.get("Assigned_User") or "Unassigned")
                users[key] = users.get(key, 0) + 1
            routing = {}
            for r in rows:
                key = str(r.get("Destination_Folder") or "Not Routed")
                routing[key] = routing.get(key, 0) + 1

            self.clear_panel(self.category_panel)
            self.clear_panel(self.user_panel)
            self.clear_panel(self.routing_panel)
            self.clear_panel(self.aging_panel)

            max_cat = max(categories.values(), default=1)
            for key, value in sorted(categories.items(), key=lambda x: (-x[1], x[0]))[:8]:
                self.add_bar_row(self.category_panel._body, key, value, max_cat)

            max_user = max(users.values(), default=1)
            for key, value in sorted(users.items(), key=lambda x: (-x[1], x[0]))[:8]:
                self.add_bar_row(self.user_panel._body, key, value, max_user)

            max_route = max(routing.values(), default=1)
            for key, value in sorted(routing.items(), key=lambda x: (-x[1], x[0]))[:8]:
                self.add_bar_row(self.routing_panel._body, key, value, max_route)

            now = datetime.now()
            buckets = {"0–4 hrs": 0, "4–8 hrs": 0, "8–24 hrs": 0, "24–48 hrs": 0, "48+ hrs": 0}
            ages = []
            for r in pending_rows:
                received = self._parse_dashboard_datetime(r.get("Received_Time"))
                if not received:
                    continue
                age_hours = max(0, (now - received).total_seconds() / 3600)
                ages.append(age_hours)
                if age_hours < 4:
                    buckets["0–4 hrs"] += 1
                elif age_hours < 8:
                    buckets["4–8 hrs"] += 1
                elif age_hours < 24:
                    buckets["8–24 hrs"] += 1
                elif age_hours < 48:
                    buckets["24–48 hrs"] += 1
                else:
                    buckets["48+ hrs"] += 1

            max_age_bucket = max(buckets.values(), default=1)
            for key, value in buckets.items():
                self.add_bar_row(self.aging_panel._body, key, value, max_age_bucket)

            if ages:
                oldest = max(ages)
                avg_age = sum(ages) / len(ages)
                critical = sum(1 for x in ages if x >= 48)
                summary = QLabel(f"Oldest pending: {self.format_age(oldest)}   •   Average: {self.format_age(avg_age)}   •   48+ hrs: {critical}")
            else:
                summary = QLabel("No pending emails with a valid Received_Time.")
            summary.setObjectName("AgingSummary")
            self.aging_panel._body.addWidget(summary)
            self.aging_panel._body.addStretch()

            self.dashboard_updated.setText("Updated " + datetime.now().strftime("%H:%M:%S"))
        except Exception as error:
            self.dashboard_updated.setText("Data refresh unavailable")

    @staticmethod
    def format_age(hours):
        if hours < 1:
            return f"{int(hours * 60)} min"
        if hours < 24:
            return f"{hours:.1f} hrs"
        return f"{hours / 24:.1f} days"

    def make_info_card(self, heading, value):
        frame = QFrame()
        frame.setObjectName("InfoCard")

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(14, 11, 14, 11)
        layout.setSpacing(4)

        h = QLabel(heading)
        h.setObjectName("CardHeading")

        v = QLabel(value)
        v.setObjectName("CardValue")
        v.setWordWrap(True)

        layout.addWidget(h)
        layout.addWidget(v)

        return frame

    def styles(self):
        return f"""
        QWidget#Root {{
            background: {self.BG};
            color: {self.TEXT};
            font-family: "Segoe UI";
        }}

        QLabel#Title {{
            font-size: 34px;
            font-weight: 900;
            color: {self.TEXT};
        }}

        QLabel#Subtitle {{
            font-size: 14px;
            font-weight: 800;
            color: {self.RED};
        }}

        QLabel#Description {{
            font-size: 12px;
            color: {self.MUTED};
        }}

        QFrame#StatusCard {{
            background: {self.PANEL};
            border: 1px solid {self.BORDER};
            border-radius: 12px;
        }}

        QLabel#StatusHeading {{
            font-size: 10px;
            font-weight: 800;
            color: {self.MUTED};
            letter-spacing: 1px;
        }}

        QLabel#StatusDot {{
            color: #6B7280;
            font-size: 17px;
        }}

        QLabel#StatusText {{
            color: #CBD1D8;
            font-size: 13px;
            font-weight: 800;
        }}

        QFrame#ControlBar {{
            background: {self.PANEL};
            border: 1px solid {self.BORDER};
            border-radius: 11px;
        }}

        QPushButton#StartButton,
        QPushButton#StopButton {{
            border-radius: 8px;
            padding: 11px 19px;
            font-size: 12px;
            font-weight: 800;
        }}

        QPushButton#StartButton {{
            background: {self.RED};
            color: white;
            border: 1px solid {self.RED};
        }}

        QPushButton#StartButton:hover {{
            background: #EF2A31;
        }}

        QPushButton#StartButton:disabled {{
            background: #642126;
            border-color: #642126;
            color: #B9A0A2;
        }}

        QPushButton#StopButton {{
            background: {self.PANEL_2};
            color: #F1B5B8;
            border: 1px solid #7D3035;
        }}

        QPushButton#StopButton:hover {{
            background: #241417;
            border-color: {self.RED};
        }}

        QPushButton#StopButton:disabled {{
            color: #656C74;
            border-color: {self.BORDER};
        }}

        QLabel#SmallInfo {{
            color: {self.MUTED};
            font-size: 10px;
            font-weight: 700;
        }}

        QFrame#InfoCard {{
            background: {self.PANEL};
            border: 1px solid {self.BORDER};
            border-radius: 9px;
        }}

        QLabel#CardHeading {{
            font-size: 9px;
            font-weight: 800;
            color: {self.MUTED};
            letter-spacing: 1px;
        }}

        QLabel#CardValue {{
            font-size: 12px;
            font-weight: 700;
            color: {self.TEXT};
        }}

        QLabel#SectionTitle {{
            font-size: 14px;
            font-weight: 800;
            color: {self.TEXT};
        }}

        QLabel#ActivityStatus {{
            font-size: 11px;
            color: {self.MUTED};
        }}

        QPlainTextEdit#Console {{
            background: {self.CONSOLE};
            color: #E8EAED;
            border: 1px solid {self.BORDER};
            border-radius: 10px;
            padding: 13px;
            font-family: "Cascadia Mono", "Consolas", monospace;
            font-size: 11px;
            selection-background-color: #343A40;
        }}

        QScrollBar:vertical {{
            background: #111418;
            width: 10px;
            margin: 2px;
            border-radius: 5px;
        }}

        QScrollBar::handle:vertical {{
            background: #343A40;
            min-height: 30px;
            border-radius: 5px;
        }}

        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {{
            height: 0px;
        }}

        QTabWidget#MainTabs::pane {{
            border: 1px solid #2A3037;
            background: #0B0D0F;
            border-radius: 10px;
            top: -1px;
        }}

        QTabBar::tab {{
            background: #121519;
            color: #9AA3AD;
            padding: 9px 18px;
            margin-right: 3px;
            border: 1px solid #2A3037;
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            font-size: 10px;
            font-weight: 800;
        }}

        QTabBar::tab:selected {{
            color: #F2F4F7;
            background: #171B20;
            border-top: 2px solid #D71920;
        }}

        QLabel#DashboardTitle {{
            font-size: 18px;
            font-weight: 900;
            color: #F2F4F7;
        }}

        QLabel#DashboardUpdated {{
            color: #7F8994;
            font-size: 10px;
        }}

        QPushButton#RefreshButton {{
            background: #171B20;
            color: #D7DCE2;
            border: 1px solid #343B43;
            border-radius: 7px;
            padding: 8px 13px;
            font-size: 10px;
            font-weight: 800;
        }}

        QPushButton#RefreshButton:hover {{
            border-color: #D71920;
            color: #FFFFFF;
        }}

        QFrame#KPICard, QFrame#DashboardPanel {{
            background: #121519;
            border: 1px solid #2A3037;
            border-radius: 10px;
        }}

        QLabel#KPIHeading, QLabel#PanelHeading {{
            color: #8F98A2;
            font-size: 9px;
            font-weight: 800;
            letter-spacing: 1px;
        }}

        QLabel#KPIValue {{
            color: #F2F4F7;
            font-size: 25px;
            font-weight: 900;
        }}

        QLabel#BarLabel {{
            color: #CBD1D8;
            font-size: 10px;
        }}

        QLabel#BarValue {{
            color: #F2F4F7;
            font-size: 10px;
            font-weight: 800;
        }}

        QProgressBar#DashboardBar {{
            background: #252B32;
            border: none;
            border-radius: 4px;
            min-height: 7px;
            max-height: 7px;
        }}

        QProgressBar#DashboardBar::chunk {{
            background: #D71920;
            border-radius: 4px;
        }}

        QLabel#AgingSummary {{
            color: #9AA3AD;
            font-size: 10px;
            padding-top: 4px;
        }}

        QLabel#Footer {{
            color: #626A73;
            font-size: 10px;
        }}
        """

    @Slot()
    def start_monitoring(self):
        if self.is_running:
            return

        self.console.clear()
        self.console.appendPlainText("=" * 75)
        self.console.appendPlainText("MAILFLOW AX DESKTOP MONITOR")
        self.console.appendPlainText("Starting local monitoring engine...")
        self.console.appendPlainText("=" * 75)

        self.is_running = True
        self.set_running_state(True)

        self.thread = QThread()
        self.worker = PipelineWorker()
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.output.connect(self.append_output)
        self.worker.finished.connect(self.pipeline_finished)
        self.worker.failed.connect(self.pipeline_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    @Slot()
    def stop_monitoring(self):
        if not self.is_running or self.worker is None:
            return

        self.activity_status.setText(
            "Stopping... finishing current operation"
        )
        self.stop_button.setEnabled(False)
        self.worker.stop()

    @Slot(str)
    def append_output(self, text):
        if text is None:
            return

        self.console.appendPlainText(text)

        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

        upper = text.upper()

        if "MAILFLOW AX IS RUNNING" in upper:
            self.activity_status.setText("Monitoring active")
        elif "NEW EMAIL DETECTED" in upper:
            self.activity_status.setText("Processing new email...")
        elif "AI CLASSIFYING" in upper:
            self.activity_status.setText("Local AI classification...")
        elif "MOVED SUCCESSFULLY" in upper:
            self.activity_status.setText("Email processed successfully")
        elif "USER ACTION DETECTED" in upper:
            self.activity_status.setText("User action detected")
        elif "AI CLASSIFICATION FAILED" in upper:
            self.activity_status.setText("AI classification error")
        elif "ALLOCATION FAILED" in upper or "MOVE FAILED" in upper:
            self.activity_status.setText("Processing error")

    @Slot()
    def pipeline_finished(self):
        self.is_running = False
        self.set_running_state(False)
        self.activity_status.setText("Monitoring stopped")
        self.worker = None
        self.thread = None

    @Slot(str)
    def pipeline_failed(self, error):
        self.append_output("\n❌ ENGINE ERROR: " + error)
        self.is_running = False
        self.set_running_state(False)
        self.activity_status.setText("Engine error")

        QMessageBox.critical(
            self,
            "MailFlow AX Error",
            "The monitoring engine stopped unexpectedly.\n\n" + error,
        )

    def set_running_state(self, running):
        if running:
            self.status_dot.setStyleSheet(
                f"color: {self.GREEN}; font-size: 17px;"
            )
            self.status_label.setText("MONITORING ACTIVE")
            self.status_label.setStyleSheet(
                f"color: {self.GREEN}; font-size: 13px; font-weight: 800"
            )
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
        else:
            self.status_dot.setStyleSheet(
                "color: #6B7280; font-size: 17px;"
            )
            self.status_label.setText("NOT RUNNING")
            self.status_label.setStyleSheet(
                "color: #CBD1D8; font-size: 13px; font-weight: 800"
            )
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)

    def closeEvent(self, event):
        self.dashboard_timer.stop()
        if self.is_running and self.worker is not None:
            answer = QMessageBox.question(
                self,
                "Stop MailFlow AX?",
                "MailFlow AX is currently monitoring the mailbox.\n\n"
                "Stop monitoring before closing the application?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if answer == QMessageBox.No:
                event.ignore()
                return

            self.worker.stop()

            if self.thread is not None:
                self.thread.quit()
                self.thread.wait(5000)

        event.accept()

def main():
    app = QApplication.instance() or QApplication([])
    app.setApplicationName("MailFlow AX")

    window = MailFlowAXWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
