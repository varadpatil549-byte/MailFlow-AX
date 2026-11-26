import os
import requests
import msal
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

if not TENANT_ID:
    raise ValueError("TENANT_ID is missing from .env")

if not CLIENT_ID:
    raise ValueError("CLIENT_ID is missing from .env")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = [
    "User.Read",
    "Mail.Read.Shared",
]

GRAPH_URL = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------
# 1. Authenticate
# ---------------------------------------------------------

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
)

flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise RuntimeError(f"Device flow failed: {flow}")

print("\n=== MailFlow AX ===\n")
print(flow["message"])
print()

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    print("❌ Authentication failed")
    print(result)
    raise SystemExit(1)

access_token = result["access_token"]

print("✅ Authentication successful")


# ---------------------------------------------------------
# 2. Read YOUR Inbox
# ---------------------------------------------------------

url = (
    f"{GRAPH_URL}/me"
    f"/mailFolders/inbox/messages"
)

params = {
    "$top": 5,
    "$select": (
        "id,subject,sender,receivedDateTime,"
        "isRead,bodyPreview"
    ),
    "$orderby": "receivedDateTime desc",
}

headers = {
    "Authorization": f"Bearer {access_token}",
    "Accept": "application/json",
}

response = requests.get(
    url,
    headers=headers,
    params=params,
    timeout=30,
)

print(f"\nGraph status: {response.status_code}")

if response.status_code != 200:
    print("❌ Could not read your mailbox.")
    print(response.text)
    raise SystemExit(1)

data = response.json()
messages = data.get("value", [])

print("✅ Personal mailbox connected")
print(f"✅ Messages returned: {len(messages)}\n")


# ---------------------------------------------------------
# 3. Display messages
# ---------------------------------------------------------

for index, message in enumerate(messages, start=1):

    sender = message.get("sender", {})

    sender_email = (
        sender
        .get("emailAddress", {})
        .get("address", "Unknown")
    )

    print(f"--- Email {index} ---")
    print(f"Subject : {message.get('subject')}")
    print(f"Sender  : {sender_email}")
    print(f"Received: {message.get('receivedDateTime')}")
    print(f"Read    : {message.get('isRead')}")
    print(f"Preview : {message.get('bodyPreview')}")
    print()