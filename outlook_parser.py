import re
from typing import Dict, Optional


def parse_floor_request(body: str) -> Optional[Dict[str, str]]:
    lines = body.splitlines()
    name = ""
    email = ""
    badge_id = ""
    badge_type = ""
    floor = ""
    project = ""

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if not name and "floor request" in stripped.lower() and "'s" in stripped:
            name = stripped.split("'s", 1)[0].strip()
        elif stripped.lower().startswith("email:"):
            email = stripped.split(":", 1)[1].strip()
        elif stripped.lower().startswith("badge id:"):
            badge_id = stripped.split(":", 1)[1].strip()
        elif stripped.lower().startswith("badge type:"):
            badge_type = stripped.split(":", 1)[1].strip()
        elif stripped.lower().startswith("floor:"):
            floor = stripped.split(":", 1)[1].strip()
        elif stripped.lower().startswith("project:"):
            project = stripped.split(":", 1)[1].strip()

    if not name and email:
        name = email.split("@", 1)[0]

    if not email:
        match = re.search(r"[\w\.-]+@[\w\.-]+", body)
        if match:
            email = match.group(0)

    if not name and not email:
        return None

    return {
        "name": name,
        "email": email,
        "badge_id": badge_id,
        "badge_type": badge_type,
        "floor": floor,
        "project": project,
    }

