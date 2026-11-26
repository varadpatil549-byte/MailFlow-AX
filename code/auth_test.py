import os
import msal
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

if not TENANT_ID or not CLIENT_ID:
    raise ValueError("TENANT_ID or CLIENT_ID is missing from .env")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = [
    "User.Read",
    "Mail.Read.Shared",
]

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
)

flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise RuntimeError(
        f"Device flow failed: {flow}"
    )

print("\n=== MailFlow AX Authentication ===\n")
print(flow["message"])
print()

result = app.acquire_token_by_device_flow(flow)

if "access_token" in result:
    print("✅ Authentication successful!")
    print("✅ Microsoft Graph access token received.")
else:
    print("❌ Authentication failed.")
    print(result)