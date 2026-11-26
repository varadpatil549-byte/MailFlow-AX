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


# ============================================================
# MAILFLOW AX
# FINAL SINGLE-PIPELINE APPLICATION
#
# Outlook -> Local AI -> Category -> Allocation -> Move
#          -> Single Excel Log -> Reply/Forward Tracking
#
# IMPORTANT RECOVERY BEHAVIOR
# ---------------------------
# At startup, AX reads ALL successfully processed Message_IDs
# from MailFlow_Log, then scans the ENTIRE Outlook Inbox.
#
# Therefore:
#   Email arrives at 03:00
#   AX is OFF
#   AX starts at 03:10
#   -> email is detected and processed.
#
# Message_ID is a stable Outlook MAPI Search Key, not EntryID.
# EntryID can change after an item is moved between folders.
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
    Return the stable Outlook MAPI Search Key.

    This is deliberately used instead of EntryID because
    EntryID may change when an Outlook item is moved.
    """

    try:
        raw_key = item.PropertyAccessor.GetProperty(
            SEARCH_KEY_SCHEMA
        )

        key = normalize_search_key(raw_key)

        if key:
            return key

    except Exception as error:
        # We do not silently fall back to EntryID because the
        # whole recovery design depends on a move-stable ID.
        raise RuntimeError(
            "Could not read Outlook PR_SEARCH_KEY for "
            "the email. AX cannot safely track this message."
        ) from error

    raise RuntimeError(
        "Outlook returned an empty PR_SEARCH_KEY."
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
    full_scan=False,
):
    """
    Collect unprocessed emails.

    Startup / recovery:
        full_scan=True
        -> entire Inbox is reconciled against Excel.

    Normal runtime:
        full_scan=False
        -> scan newest-first and stop once the first known
           processed ID is reached.
    """

    items = inbox.Items

    items.Sort(
        "[ReceivedTime]",
        True,
    )

    candidates = []

    scanned = 0

    for item in items:

        if (
            not full_scan
            and scanned >= MAX_RUNTIME_SCAN
        ):
            break

        if not is_mail(item):
            continue

        scanned += 1

        try:
            message_id = get_message_id(
                item
            )
        except Exception as error:
            print(
                f"⚠️ Could not identify an Inbox email: "
                f"{error}"
            )
            continue

        if message_id in processed_ids:

            # During normal runtime, all newer items have
            # already been considered. This makes polling
            # much faster than scanning the whole Inbox every
            # two seconds.
            if not full_scan:
                break

            continue

        candidates.append(
            item
        )

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
# MAIN
# ============================================================

def main():
    print("\n" + "=" * 75)
    print(
        "MAILFLOW AX - FINAL SINGLE PIPELINE"
    )
    print(
        "OUTLOOK -> AI -> ALLOCATION -> MOVE -> ACTION"
    )
    print("=" * 75)

    create_final_workbook()
    validate_final_workbook()

    config = load_client_config()

    pythoncom.CoInitialize()

    try:
        namespace = connect_outlook()

        inbox, sent_folder = (
            get_mail_folders(
                namespace
            )
        )

        print(
            f"✅ Monitoring Inbox: "
            f"{inbox.Name}"
        )

        print(
            f"✅ Monitoring Sent Items: "
            f"{sent_folder.Name}"
        )

        # ----------------------------------------------------
        # STARTUP / RECOVERY RECONCILIATION
        #
        # This is intentionally a FULL Inbox scan.
        # ----------------------------------------------------

        processed_ids = (
            load_processed_message_ids()
        )

        print(
            f"✅ Successfully processed IDs "
            f"loaded from Excel: "
            f"{len(processed_ids)}"
        )

        startup_candidates = (
            get_inbox_candidates(
                inbox,
                processed_ids,
                full_scan=True,
            )
        )

        print(
            f"🔎 Startup reconciliation found "
            f"{len(startup_candidates)} "
            f"unprocessed Inbox email(s)."
        )

        # Process every email that is in Outlook but not yet
        # successfully recorded as Moved in Excel.
        for item in startup_candidates:

            success = process_new_email(
                item,
                inbox,
                config,
            )

            if success:

                try:
                    processed_ids.add(
                        get_message_id(item)
                    )
                except Exception:
                    # The item has already moved; its stable key
                    # was written to Excel.
                    pass

        # ----------------------------------------------------
        # ACTION TRACKING RECOVERY
        # ----------------------------------------------------

        action_records = (
            load_open_action_records()
        )

        refresh_action_records(
            inbox,
            action_records,
        )

        seen_sent_ids = (
            build_seen_sent_ids(
                sent_folder
            )
        )

        print(
            f"✅ Open action records: "
            f"{len(action_records)}"
        )

        print(
            "\n🟢 MAILFLOW AX IS RUNNING"
        )

        print(
            "New Email -> AI -> Allocation -> Move -> Action"
        )

        print(
            "Startup recovery: FULL Inbox reconciliation"
        )

        print(
            "Normal polling: newest Inbox items"
        )

        print(
            "Actioned = Reply / Reply All / Forward ONLY"
        )

        print(
            "Open/Read = ignored"
        )

        print(
            "Category/Flag = logged only"
        )

        print(
            "\nPress Ctrl+C to stop."
        )

        print("=" * 75)

        # ----------------------------------------------------
        # NORMAL RUNTIME
        # ----------------------------------------------------

        while True:

            # Refresh Client_Config so routing can be edited
            # in Excel between tests without restarting AX.
            config = load_client_config()

            # ------------------------------------------------
            # A. NEW EMAILS
            # ------------------------------------------------

            runtime_candidates = (
                get_inbox_candidates(
                    inbox,
                    processed_ids,
                    full_scan=False,
                )
            )

            for item in runtime_candidates:

                try:
                    message_id = (
                        get_message_id(item)
                    )
                except Exception:
                    continue

                success = process_new_email(
                    item,
                    inbox,
                    config,
                )

                if success:
                    processed_ids.add(
                        message_id
                    )

            # ------------------------------------------------
            # B. CATEGORY / FLAG
            # ------------------------------------------------

            action_records = (
                load_open_action_records()
            )

            refresh_action_records(
                inbox,
                action_records,
            )

            check_category_and_flag_changes(
                inbox,
                action_records,
            )

            # ------------------------------------------------
            # C. REPLY / REPLY ALL / FORWARD
            # ------------------------------------------------

            check_new_sent_items(
                sent_folder,
                action_records,
                seen_sent_ids,
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
