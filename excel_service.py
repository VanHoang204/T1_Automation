import os
from datetime import datetime
from typing import List, Dict, Any

import pandas as pd
from openpyxl import Workbook


SHEET_NAME = "Requests"
COLUMNS = [
    "ID",
    "CustomerName",
    "Email",
    "BadgeID",
    "BadgeType",
    "Floor",
    "Project",
    "Status",
    "CreatedTime",
]


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_DIR = os.path.join(BASE_DIR, "excel_storage")
EXCEL_FILE = os.path.join(EXCEL_DIR, "data_requests.xlsx")


def load_excel() -> None:
    os.makedirs(EXCEL_DIR, exist_ok=True)

    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME

        ws.append(COLUMNS)

        demo_rows = [
            [
                1,
                "Erik Martin",
                "erimart@microsoft.com",
                "VU2600046",
                "Temporary",
                "F7 4F-ODT Area",
                "",
                "Approved",
                datetime(2026, 1, 1),
            ],
            [
                2,
                "Jonathan Hutchinson",
                "jonathan.hutchinson@microsoft.com",
                "VU2600047",
                "Temporary",
                "F3 2F-Open Area",
                "",
                "Approved",
                datetime(2026, 1, 2),
            ],
            [
                3,
                "David Lee",
                "david.lee@microsoft.com",
                "VU2600048",
                "Visitor",
                "F5 3F-Desk Area",
                "",
                "Pending",
                datetime(2026, 1, 3),
            ],
            [
                4,
                "Anna Smith",
                "anna.smith@microsoft.com",
                "VU2600049",
                "Temporary",
                "F7 4F-ODT Area",
                "",
                "Approved",
                datetime(2026, 1, 4),
            ],
        ]

        for row in demo_rows:
            ws.append(row)

        wb.save(EXCEL_FILE)


def _read_dataframe() -> pd.DataFrame:
    load_excel()
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine="openpyxl")

    if df.empty:
        df = pd.DataFrame(columns=COLUMNS)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df


def _write_dataframe(df: pd.DataFrame) -> None:
    df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False, engine="openpyxl")


def read_all() -> List[Dict[str, Any]]:
    df = _read_dataframe()

    if df.empty:
        return []

    df = df.fillna("")

    records: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        record = {col: row.get(col, "") for col in COLUMNS}
        try:
            record["ID"] = int(record.get("ID", 0))
        except (ValueError, TypeError):
            record["ID"] = 0
        records.append(record)

    return records


def _next_id(df: pd.DataFrame) -> int:
    if df.empty:
        return 1

    try:
        max_id = int(df["ID"].max())
        return max_id + 1
    except Exception:
        return 1


def add_record(data: Dict[str, Any]) -> int:
    df = _read_dataframe()
    new_id = _next_id(df)

    new_row = {
        "ID": new_id,
        "CustomerName": data.get("CustomerName", ""),
        "Email": data.get("Email", ""),
        "BadgeID": data.get("BadgeID", ""),
        "BadgeType": data.get("BadgeType", ""),
        "Floor": data.get("Floor", ""),
        "Project": data.get("Project", ""),
        "Status": data.get("Status", ""),
        "CreatedTime": datetime.now(),
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    _write_dataframe(df)

    return new_id


def update_record(record_id: int, data: Dict[str, Any]) -> bool:
    df = _read_dataframe()

    if df.empty:
        return False

    try:
        df["ID"] = df["ID"].astype(int)
    except Exception:
        return False

    mask = df["ID"] == int(record_id)
    if not mask.any():
        return False

    for key in ["CustomerName", "Email", "BadgeID", "BadgeType", "Floor", "Project", "Status"]:
        if key in data:
            df.loc[mask, key] = data[key]

    _write_dataframe(df)
    return True


def delete_record(record_id: int) -> bool:
    df = _read_dataframe()

    if df.empty:
        return False

    try:
        df["ID"] = df["ID"].astype(int)
    except Exception:
        return False

    before_count = len(df)
    df = df[df["ID"] != int(record_id)]

    if len(df) == before_count:
        return False

    _write_dataframe(df)
    return True


def get_excel_file_path() -> str:
    load_excel()
    return EXCEL_FILE
