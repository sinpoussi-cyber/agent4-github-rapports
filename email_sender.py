import os
import base64
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

TOKEN_FILE = "token.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


def _load_credentials() -> Credentials | None:
    """Charge et rafraîchit les credentials OAuth depuis token.json."""
    if not os.path.exists(TOKEN_FILE):
        print(f"[email_sender] Fichier {TOKEN_FILE} introuvable.")
        return None

    try:
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            token_data = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"[email_sender] Impossible de lire {TOKEN_FILE}: {e}")
        return None

    creds = Credentials(
        token=token_data.get("token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=token_data.get("client_id") or os.environ.get("GMAIL_CLIENT_ID", ""),
        client_secret=token_data.get("client_secret") or os.environ.get("GMAIL_CLIENT_SECRET", ""),
        scopes=token_data.get("scopes", SCOPES),
    )

    # Rafraîchit le token si expiré
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            # Persiste le token rafraîchi
            updated = json.loads(creds.to_json())
            with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                json.dump(updated, f, indent=2)
        except Exception as e:
            print(f"[email_sender] Échec du rafraîchissement du token: {e}")
            return None

    return creds


def _build_message(sender: str, recipient: str,
                   subject: str, body_html: str, body_text: str) -> dict:
    """Construit un email multipart/alternative encodé en base64url."""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient

    msg.attach(MIMEText(body_text, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return {"raw": raw}


def send_report(subject: str, body_html: str, body_text: str) -> bool:
    """
    Envoie le rapport par email via l'API Gmail.

    Returns:
        True si l'email a été envoyé, False sinon.
    """
    sender = os.environ.get("REPORT_EMAIL_FROM", "")
    recipient = os.environ.get("REPORT_EMAIL_TO", "")

    if not sender or not recipient:
        print("[email_sender] REPORT_EMAIL_FROM et REPORT_EMAIL_TO doivent être définis dans .env")
        return False

    creds = _load_credentials()
    if creds is None:
        return False

    try:
        service = build("gmail", "v1", credentials=creds, cache_discovery=False)
        message = _build_message(sender, recipient, subject, body_html, body_text)
        sent = service.users().messages().send(userId="me", body=message).execute()
        print(f"[email_sender] Email envoyé avec succès. Message ID: {sent.get('id')}")
        print(f"[email_sender] De : {sender} → À : {recipient}")
        print(f"[email_sender] Objet : {subject}")
        return True

    except HttpError as e:
        print(f"[email_sender] Erreur Gmail API ({e.resp.status}): {e.error_details}")
        return False
    except Exception as e:
        print(f"[email_sender] Erreur inattendue: {e}")
        return False


if __name__ == "__main__":
    ok = send_report(
        subject="Test Agent GitHub Rapports",
        body_html="<h1>Test</h1><p>Ceci est un email de test de l'Agent GitHub Rapports.</p>",
        body_text="Test - Ceci est un email de test de l'Agent GitHub Rapports.",
    )
    print("Résultat :", "OK" if ok else "ÉCHEC")
