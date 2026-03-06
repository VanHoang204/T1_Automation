import logging
import os
from datetime import datetime

import win32com.client

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "DjangoProject1.settings")

import django  # noqa: E402

django.setup()

from requests_app.models import FloorRequest  # noqa: E402
from outlook_parser import parse_floor_request  # noqa: E402


LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
LOG_FILE = os.path.join(LOG_DIR, "outlook_listener.log")

os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


def backfill_inbox() -> None:
    logging.info("Starting Outlook backfill")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
    except Exception as exc:
        logging.error("Failed to access Outlook Inbox: %s", exc)
        return

    count = 0

    for item in items:
        try:
            sender = str(getattr(item, "SenderEmailAddress", "") or "")
            subject = str(getattr(item, "Subject", "") or "")
            body = str(getattr(item, "Body", "") or "")
            entry_id = str(getattr(item, "EntryID", "") or "")
        except Exception:
            continue

        if not entry_id:
            continue

        if "t1approval@microsoft.com" not in sender.lower() and "floor request" not in subject.lower():
            continue

        if FloorRequest.objects.filter(mail_entry_id=entry_id).exists():
            continue

        parsed = parse_floor_request(body)
        if not parsed:
            logging.warning("Could not parse email body for backfill entry_id=%s", entry_id)
            continue

        FloorRequest.objects.create(
            name=parsed.get("name", ""),
            email=parsed.get("email", ""),
            badge_id=parsed.get("badge_id", ""),
            badge_type=parsed.get("badge_type", ""),
            floor=parsed.get("floor", ""),
            project=parsed.get("project", ""),
            mail_subject=subject,
            mail_entry_id=entry_id,
            request_time=datetime.now(),
            is_read=False,
            is_processed=False,
        )
        count += 1

    logging.info("Backfill completed, created %s records", count)


if __name__ == "__main__":
    backfill_inbox()

