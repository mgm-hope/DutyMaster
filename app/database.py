from __future__ import annotations

import sqlite3
from pathlib import Path

from .constants import DATA_DIR, DB_PATH


SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS timetable_meta (
    id INTEGER PRIMARY KEY,
    school_name TEXT,
    week1_label TEXT,
    week2_label TEXT,
    uploaded_at TEXT
);

CREATE TABLE IF NOT EXISTS teachers (
    initials TEXT PRIMARY KEY,
    full_name TEXT,
    is_teaching INTEGER DEFAULT 1,
    lessons_week1 REAL DEFAULT 0,
    lessons_week2 REAL DEFAULT 0,
    total_lessons REAL DEFAULT 0,
    non_contact REAL DEFAULT 0,
    protected_periods INTEGER DEFAULT 6,
    classification TEXT DEFAULT 'Teacher',
    is_part_time INTEGER DEFAULT 0,
    days_in_school TEXT DEFAULT '1111111111',
    last_updated TEXT
);

CREATE TABLE IF NOT EXISTS teacher_periods (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    teacher_initials TEXT NOT NULL,
    week INTEGER NOT NULL CHECK(week IN (1,2)),
    day TEXT NOT NULL,
    period TEXT NOT NULL,
    source_row INTEGER,
    source_col INTEGER,
    UNIQUE(teacher_initials, week, day, period)
);

CREATE TABLE IF NOT EXISTS staff_names (
    initials TEXT PRIMARY KEY,
    full_name TEXT NOT NULL,
    last_updated TEXT
);

CREATE TABLE IF NOT EXISTS classifications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS additional_staff (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category TEXT NOT NULL CHECK(category IN ('Pastoral','Admin','ESLT','Chaplaincy','SLT')),
    initials TEXT UNIQUE NOT NULL,
    full_name TEXT NOT NULL,
    is_full_time INTEGER DEFAULT 1,
    days_in_school TEXT DEFAULT '1111111111',
    availability TEXT,
    is_archived INTEGER DEFAULT 0,
    status TEXT DEFAULT 'Active',
    last_updated TEXT
);

CREATE TABLE IF NOT EXISTS rota_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    week INTEGER NOT NULL CHECK(week IN (1,2)),
    day TEXT NOT NULL,
    period TEXT NOT NULL,
    staff_type TEXT,
    staff_initials TEXT,
    last_updated TEXT,
    UNIQUE(week, day, period)
);

CREATE TABLE IF NOT EXISTS rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    description TEXT,
    active INTEGER DEFAULT 1,
    last_updated TEXT
);

CREATE TABLE IF NOT EXISTS problem_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
    issue_type TEXT,
    description TEXT,
    week INTEGER,
    day TEXT,
    period TEXT
);
"""


DEFAULT_RULES = [
    ("Gate Duty = SLT Only", "Gate Duty (08:00) can only be assigned to SLT members"),
    ("Monday P6 = AOW", "Monday Period 6 is automatically AOW for all staff unless they teach P6"),
    ("Tutor First Duty = SLT", "First duty during Tutor Time must be SLT"),
    ("Pastoral Roles", "Pastoral Support, Room 90, and Isolation prefer Pastoral staff then SLT"),
    ("Period 4 Lunch Rules", "7 staff total, 1 Pastoral minimum, max 2 SLT, no P4 teaching clash"),
    ("Period 4 Mutual Exclusion", "Staff cannot be on P4 lunch and 4A/4B/4C duties on the same day"),
    ("Period 7 Detention = 2 Staff", "Exactly 2 staff required for detention duty"),
    ("Part-Time Day Protection", "Do not assign staff on days they are out of school"),
    ("Respect Protected Periods", "Keep each teacher's protected periods free where possible"),
]


def get_connection(path: Path | None = None) -> sqlite3.Connection:
    target = Path(path or DB_PATH)
    target.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(target)
    conn.row_factory = sqlite3.Row
    return conn


def initialise_database(path: Path | None = None) -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = get_connection(path)
    conn.executescript(SCHEMA_SQL)
    migrate_database(conn)
    seed_defaults(conn)
    conn.commit()
    return conn


def migrate_database(conn: sqlite3.Connection) -> None:
    _ensure_column(conn, "teachers", "days_in_school", "TEXT DEFAULT '1111111111'")
    _ensure_column(conn, "additional_staff", "days_in_school", "TEXT DEFAULT '1111111111'")
    _ensure_column(conn, "additional_staff", "is_archived", "INTEGER DEFAULT 0")
    _ensure_column(conn, "additional_staff", "status", "TEXT DEFAULT 'Active'")


def _ensure_column(conn: sqlite3.Connection, table_name: str, column_name: str, column_type: str) -> None:
    columns = {row["name"] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}
    if column_name not in columns:
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")


def seed_defaults(conn: sqlite3.Connection) -> None:
    conn.executemany(
        "INSERT OR IGNORE INTO classifications(name) VALUES (?)",
        [("Teacher",), ("HOF",), ("SLT",)],
    )
    conn.executemany(
        "INSERT OR IGNORE INTO rules(name, description) VALUES (?, ?)",
        DEFAULT_RULES,
    )
