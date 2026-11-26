from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_FOLDER = PROJECT_ROOT / "data"
EXCEL_FILE = DATA_FOLDER / "MailFlow_AX.xlsx"


def create_excel_workbook():
    DATA_FOLDER.mkdir(exist_ok=True)

    # Don't overwrite an existing workbook.
    if EXCEL_FILE.exists():
        print(f"Excel file already exists: {EXCEL_FILE}")
        return

    workbook = Workbook()

    # -------------------------------------------------
    # Email Log
    # -------------------------------------------------

    email_log = workbook.active
    email_log.title = "Email_Log"

    headers = [
        "Ticket_ID",
        "Message_ID",
        "Received_Time",
        "Sender_Name",
        "Sender_Email",
        "Subject",
        "Body_Preview",
        "AI_Category",
        "AI_Confidence",
        "Captured_Time",
        "Status",
    ]

    email_log.append(headers)

    # Header formatting
    for cell in email_log[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Column widths
    widths = {
        "A": 15,
        "B": 45,
        "C": 25,
        "D": 25,
        "E": 35,
        "F": 50,
        "G": 80,
        "H": 25,
        "I": 18,
        "J": 25,
        "K": 20,
    }

    for column, width in widths.items():
        email_log.column_dimensions[column].width = width

    email_log.freeze_panes = "A2"

    # -------------------------------------------------
    # System Log
    # -------------------------------------------------

    system_log = workbook.create_sheet("System_Log")

    system_headers = [
        "Timestamp",
        "Module",
        "Level",
        "Message",
    ]

    system_log.append(system_headers)

    for cell in system_log[1]:
        cell.font = Font(bold=True)

    system_log.column_dimensions["A"].width = 25
    system_log.column_dimensions["B"].width = 25
    system_log.column_dimensions["C"].width = 15
    system_log.column_dimensions["D"].width = 100

    # -------------------------------------------------
    # Configuration
    # -------------------------------------------------

    config = workbook.create_sheet("Configuration")

    config_headers = [
        "Parameter",
        "Value",
    ]

    config.append(config_headers)

    for cell in config[1]:
        cell.font = Font(bold=True)

    config.append(["Project_Name", "MailFlow AX"])
    config.append(["Version", "Phase 1"])
    config.append(["Mailbox_Connector", "Outlook COM"])
    config.append(["AI_Status", "Not configured"])
    config.append(["Allocation_Status", "Not started"])
    config.append(["Action_Tracking", "Not started"])

    config.column_dimensions["A"].width = 30
    config.column_dimensions["B"].width = 40

    workbook.save(EXCEL_FILE)

    print("✅ MailFlow AX Excel workbook created")
    print(f"📁 {EXCEL_FILE}")


if __name__ == "__main__":
    create_excel_workbook()