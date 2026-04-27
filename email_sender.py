import smtplib
import os
import logging
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

logger = logging.getLogger(__name__)


def send_report(
    subject: str,
    html_body: str,
    body_text: str = "",
    attachments: list = None,
) -> bool:
    """
    Envoie un email HTML avec pièces jointes optionnelles.
    attachments : liste de dicts {"filename": str, "data": bytes}
    """
    sender      = os.environ.get("REPORT_EMAIL_FROM", "")
    recipient   = os.environ.get("REPORT_EMAIL_TO", "")
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    port_str    = os.environ.get("SMTP_PORT", "587")
    port        = int(port_str.strip()) if port_str and port_str.strip().isdigit() else 587
    password    = os.environ.get("SMTP_PASSWORD", "")

    if not all([sender, recipient, password]):
        logger.error("Missing SMTP credentials")
        return False

    if attachments:
        msg = MIMEMultipart("mixed")
        msg["Subject"] = subject
        msg["From"]    = sender
        msg["To"]      = recipient
        alt = MIMEMultipart("alternative")
        if body_text:
            alt.attach(MIMEText(body_text, "plain"))
        alt.attach(MIMEText(html_body, "html"))
        msg.attach(alt)
        for att in attachments:
            part = MIMEApplication(att["data"], Name=att["filename"])
            part["Content-Disposition"] = f'attachment; filename="{att["filename"]}"'
            msg.attach(part)
    else:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = sender
        msg["To"]      = recipient
        msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP(smtp_server, port, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender, password)
            smtp.sendmail(sender, recipient, msg.as_string())
        logger.info(f"Email envoyé à {recipient}")
        return True
    except Exception as e:
        logger.error(f"SMTP error: {e}")
        return False
