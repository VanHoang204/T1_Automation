import logging
import os
import time
from datetime import datetime
from typing import Optional

import pythoncom
import win32com.client
from django.db import IntegrityError

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


outlook_namespace = None


def _process_entry(entry_id: str) -> None:
    global outlook_namespace
    try:
        mail = outlook_namespace.GetItemFromID(entry_id)
    except Exception as exc:
        logging.error("Failed to get mail item %s: %s", entry_id, exc)
        return

    try:
        sender = str(getattr(mail, "SenderEmailAddress", "") or "")
        subject = str(getattr(mail, "Subject", "") or "")
        body = str(getattr(mail, "Body", "") or "")
    except Exception as exc:
        logging.error("Failed to read properties for %s: %s", entry_id, exc)
        return

    logging.info("Email received: entry_id=%s subject=%s", entry_id, subject)

    if "t1approval@microsoft.com" not in sender.lower() and "floor request" not in subject.lower():
        return

    parsed = parse_floor_request(body)
    if not parsed:
        logging.warning("Could not parse email body for entry_id=%s", entry_id)
        return

    try:
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
        )
        logging.info("Record saved for entry_id=%s", entry_id)
    except IntegrityError:
        logging.info("Duplicate entry_id ignored: %s", entry_id)
    except Exception as exc:
        logging.error("Failed to save record for %s: %s", entry_id, exc)


class OutlookEventHandler:
    def OnNewMailEx(self, entry_id_collection: str) -> None:
        if not entry_id_collection:
            return
        for entry_id in entry_id_collection.split(","):
            _process_entry(entry_id)


def run_listener() -> None:
    global outlook_namespace

    logging.info("Starting Outlook listener")
    try:
        outlook = win32com.client.DispatchWithEvents(
            "Outlook.Application", OutlookEventHandler
        )
        outlook_namespace = outlook.GetNamespace("MAPI")
    except Exception as exc:
        logging.error("Failed to connect to Outlook: %s", exc)
        return

    logging.info("Outlook listener connected, waiting for NewMailEx events")
    while True:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.5)


if __name__ == "__main__":
    run_listener()

