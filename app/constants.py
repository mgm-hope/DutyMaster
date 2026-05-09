from __future__ import annotations

import os
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = Path(os.getenv("DUTYMASTER_DATA_DIR", "/data" if Path("/data").exists() else str(APP_DIR / "data")))
DB_PATH = Path(os.getenv("DUTYMASTER_DB_PATH", str(DATA_DIR / "dutymaster_online.sqlite3")))
UPLOADS_DIR = Path(os.getenv("DUTYMASTER_UPLOAD_DIR", str(DATA_DIR / "uploads")))
EXPORTS_DIR = Path(os.getenv("DUTYMASTER_EXPORT_DIR", str(DATA_DIR / "exports")))
MASTER_TEMPLATE_PATH = Path(os.getenv("DUTYMASTER_MASTER_TEMPLATE", str(DATA_DIR / "2025 MASTER DUTY FINAL.xlsx")))

ROTA_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]
ROTA_WEEKS = [1, 2]

DUTY_SECTIONS = [
    ("Gate Duty", [("Gate", "Gate Duty")]),
    ("Tutor Time", [
        ("Tutor_1st_Duty", "1st Duty"),
        ("Tutor_AOW", "Act of Worship"),
        ("Tutor_Pastoral_Support", "Pastoral Support"),
        ("Tutor_Room_90", "Room 90"),
        ("Tutor_Isolation", "Isolation"),
    ]),
    ("Period 1", [
        ("P1_First_Duty", "First Duty"),
        ("P1_Pastoral_Support", "Pastoral Support"),
        ("P1_Room_90", "Room 90"),
        ("P1_Isolation", "Isolation"),
    ]),
    ("Period 2", [
        ("P2_First_Duty", "First Duty"),
        ("P2_Pastoral_Support", "Pastoral Support"),
        ("P2_Room_90", "Room 90"),
        ("P2_Isolation", "Isolation"),
    ]),
    ("Break", [
        ("Break_Duty_Lead", "Duty Lead"),
        ("Break_Late_Detention", "Late Detention"),
        ("Break_Pastoral_Support", "Pastoral Support"),
        ("Break_Room_90", "Room 90"),
        ("Break_Isolation", "Isolation"),
    ]),
    ("Teaching Staff Break Rota", [
        ("Teacher_Break_Rota_1", "Teaching Staff Break Rota 1"),
        ("Teacher_Break_Rota_2", "Teaching Staff Break Rota 2"),
        ("Teacher_Break_Rota_3", "Teaching Staff Break Rota 3"),
        ("Teacher_Break_Rota_4", "Teaching Staff Break Rota 4"),
        ("Teacher_Break_Rota_5", "Teaching Staff Break Rota 5"),
        ("Teacher_Break_Rota_6", "Teaching Staff Break Rota 6"),
        ("Teacher_Break_Rota_7", "Teaching Staff Break Rota 7"),
        ("Teacher_Break_Rota_8", "Teaching Staff Break Rota 8"),
        ("Teacher_Break_Rota_9", "Teaching Staff Break Rota 9"),
        ("Teacher_Break_Rota_10", "Teaching Staff Break Rota 10"),
    ]),
    ("Period 3", [
        ("P3_First_Duty", "First Duty"),
        ("P3_Pastoral_Support", "Pastoral Support"),
        ("P3_Room_90", "Room 90"),
        ("P3_Isolation", "Isolation"),
    ]),
    ("Period 4 Lunch", [
        ("P4_Lunch_2", "Lunch Duty 2"),
        ("P4_Lunch_3", "Lunch Duty 3"),
        ("P4_Lunch_4", "Lunch Duty 4"),
        ("P4_Lunch_5", "Lunch Duty 5"),
        ("P4_Lunch_6", "Lunch Duty 6"),
    ]),
    ("Period 4A", [
        ("P4A_First_Duty", "First Duty"),
        ("P4A_Lunch_Duty", "Lunch Duty"),
        ("P4A_Pastoral_Support", "Pastoral Support"),
        ("P4A_Isolation", "Isolation"),
        ("P4A_Rest_Break", "Rest Break"),
    ]),
    ("Period 4B", [
        ("P4B_First_Duty", "First Duty"),
        ("P4B_Lunch_Duty", "Lunch Duty"),
        ("P4B_Pastoral_Support", "Pastoral Support"),
        ("P4B_Isolation", "Isolation"),
        ("P4B_Rest_Break", "Rest Break"),
    ]),
    ("Period 4C", [
        ("P4C_First_Duty", "First Duty"),
        ("P4C_Lunch_Duty", "Lunch Duty"),
        ("P4C_Pastoral_Support", "Pastoral Support"),
        ("P4C_Isolation", "Isolation"),
        ("P4C_Rest_Break", "Rest Break"),
    ]),
    ("Period 5", [
        ("P5_First_Duty", "First Duty"),
        ("P5_Pastoral_Support", "Pastoral Support"),
        ("P5_Room_90", "Room 90"),
        ("P5_Isolation", "Isolation"),
    ]),
    ("Period 6", [
        ("P6_First_Duty", "First Duty"),
        ("P6_Pastoral_Support", "Pastoral Support"),
        ("P6_Room_90", "Room 90"),
        ("P6_Isolation", "Isolation"),
    ]),
    ("Period 7", [
        ("P7_Homework_Club", "Homework Club"),
        ("P7_Detention_1", "Detention Duty 1"),
        ("P7_Detention_2", "Detention Duty 2"),
    ]),
]

DUTY_LABELS = {code: label for _, events in DUTY_SECTIONS for code, label in events}
DUTY_ORDER = {code: index for index, code in enumerate(DUTY_LABELS)}
