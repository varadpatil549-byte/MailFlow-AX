import win32com.client


# Outlook folder constant:
# 6 = Inbox
OL_FOLDER_INBOX = 6

# Outlook item class:
# 43 = MailItem
OL_CLASS_MAIL = 43


def get_outlook():
    """Connect to the locally installed Outlook desktop application."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as exc:
        raise RuntimeError(
            "Could not connect to Outlook Desktop. "
            "Make sure Outlook is installed and your Outlook profile is configured."
        ) from exc


def main():
    print("\n=== MailFlow AX - Outlook Connection Test ===\n")

    outlook, namespace = get_outlook()

    # Confirm the Outlook profile/account currently being used.
    current_user = namespace.CurrentUser
    print(f"Outlook profile user : {current_user.Name}")

    # Get the default Inbox of the current Outlook profile.
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)

    print(f"Inbox folder         : {inbox.Name}")
    print(f"Total inbox items    : {inbox.Items.Count}")

    # Sort newest first.
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    print("\nLatest emails:\n")

    displayed = 0

    for item in items:
        # Ignore calendar items, meeting requests, reports, etc.
        if getattr(item, "Class", None) != OL_CLASS_MAIL:
            continue

        displayed += 1

        subject = item.Subject or "(No subject)"
        sender = getattr(item, "SenderName", "") or "(Unknown sender)"

        try:
            sender_email = item.SenderEmailAddress or ""
        except Exception:
            sender_email = ""

        received = item.ReceivedTime
        unread = getattr(item, "UnRead", False)

        print(f"{displayed}. {subject}")
        print(f"   Sender   : {sender}")
        print(f"   Email    : {sender_email}")
        print(f"   Received : {received}")
        print(f"   Unread   : {unread}")
        print()

        if displayed >= 5:
            break

    if displayed == 0:
        print("No mail items were found in the Inbox.")

    print("✅ Outlook connection test completed successfully.")


if __name__ == "__main__":
    main()
