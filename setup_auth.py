import os
import json
from pathlib import Path

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
except ImportError:
    raise SystemExit(
        "Missing auth dependencies. Install with:\n"
        "  python -m pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client"
    )

BASE_DIR = Path(__file__).resolve().parent
TOKEN_PATH = BASE_DIR / "token.json"
SECRET_PATH = BASE_DIR / "client_secret.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


def main():
    if not SECRET_PATH.exists():
        raise SystemExit(
            f"Missing {SECRET_PATH.name}.\n"
            "Follow the setup guide to download it from Google Cloud Console."
        )

    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(SECRET_PATH), SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w") as f:
            json.dump(json.loads(creds.to_json()), f, indent=2)

    print("Authentication successful.")
    print(f"Token saved to: {TOKEN_PATH.name}")
    print("You can now run the main parser script.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nAuth error: {exc}")
        raise SystemExit(1)