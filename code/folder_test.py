import pythoncom
import win32com.client


INBOX_FOLDER = 6


def main():

    print("\n=== MailFlow AX - Phase 2 Folder Test ===\n")

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.GetActiveObject(
            "Outlook.Application"
        )

        namespace = outlook.GetNamespace("MAPI")

        inbox = namespace.GetDefaultFolder(
            INBOX_FOLDER
        )

        print(f"✅ Inbox found: {inbox.Name}")
        print("\nChild folders:\n")

        for folder in inbox.Folders:
            print(
                f"Folder: {folder.Name} | "
                f"Items: {folder.Items.Count}"
            )

        print("\n✅ Folder discovery completed.")

    except Exception as error:
        print(f"❌ Error: {error}")

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()