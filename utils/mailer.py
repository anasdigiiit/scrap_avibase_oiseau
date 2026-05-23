from __future__ import annotations

import logging
import mimetypes
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, Optional


SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465


def email_is_configured() -> bool:
    return all(os.getenv(key) for key in ("EMAIL_USER", "EMAIL_PASSWORD", "EMAIL_TO"))


def send_email(
    *,
    subject: str,
    body: str,
    attachments: Optional[Iterable[str]] = None,
    logger: Optional[logging.Logger] = None,
) -> bool:
    email_user = os.getenv("EMAIL_USER")
    email_password = os.getenv("EMAIL_PASSWORD")
    email_to = os.getenv("EMAIL_TO")

    if not all((email_user, email_password, email_to)):
        if logger:
            logger.warning("Email notification skipped: missing EMAIL_USER, EMAIL_PASSWORD or EMAIL_TO.")
        return False

    message = EmailMessage()
    message["From"] = email_user
    message["To"] = email_to
    message["Subject"] = subject
    message.set_content(body)

    for attachment in attachments or []:
        file_path = Path(attachment)
        if not file_path.exists():
            if logger:
                logger.warning("Email attachment not found: %s", file_path)
            continue

        mime_type, _ = mimetypes.guess_type(file_path.name)
        maintype, subtype = (mime_type or "application/octet-stream").split("/", 1)
        with file_path.open("rb") as handle:
            message.add_attachment(
                handle.read(),
                maintype=maintype,
                subtype=subtype,
                filename=file_path.name,
            )

    recipients = [item.strip() for item in email_to.split(",") if item.strip()]

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.login(email_user, email_password)
            smtp.send_message(message, to_addrs=recipients)
    except Exception as exc:
        if logger:
            logger.error("Unable to send email notification: %s", exc, exc_info=True)
        return False

    if logger:
        logger.info("Email notification sent to %s", ", ".join(recipients))
    return True
