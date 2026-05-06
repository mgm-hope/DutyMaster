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

CREATE TABLE IF NOT EXISTS custom_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    active INTEGER DEFAULT 1,
    duty_scope TEXT DEFAULT 'Any',
    staff_scope TEXT DEFAULT 'Any',
    condition_type TEXT NOT NULL,
    condition_value TEXT,
    priority TEXT DEFAULT 'Hard',
    notes TEXT,
    is_archived INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    last_updated TEXT
);

CREATE TABLE IF NOT EXISTS versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    reason TEXT,
    snapshot_json TEXT NOT NULL,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
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
    ("Gate Duty = SLT Only", "Gate Duty (08:00) can only be assigned to SLT members."),
    ("Monday P6 = AOW", "Monday Period 6 is automatically Act of Worship for all staff unless they have a real class recorded in the imported timetable. AOW counts as teaching time."),
    ("Tutor First Duty = SLT", "First duty during Tutor Time must be SLT and must be distinct from assembly/AOW staff."),
    ("Tutor Pastoral Layers = Pastoral", "Pastoral Support, Room 90, and Isolation during Tutor Time are staffed by Pastoral department members where possible."),
    ("Period 4 Lunch Rules", "Lunch duty requires 7 staff total, at least 1 Pastoral member, ideally no more than 2 SLT, excludes staff teaching Period 4, and respects protected periods and part-time days."),
    ("Period 7 Detention = 2 Staff", "Period 7 detention duty requires exactly 2 staff."),
    ("Break Duty Lead = SLT or Pastoral", "Break Duty Lead can only be assigned to SLT or Pastoral staff."),
    ("Even SLT Isolation Distribution", "SLT isolation duties should be spread as evenly as possible across participating SLT members."),
    ("Even Pastoral Distribution", "Pastoral staff should be spread evenly across Pastoral Support, Room 90, Isolation, late detention, and lunch pastoral duties."),
    ("No Double Booking", "Staff cannot be assigned to two duties in the same time slot."),
    ("Part-Time Day Protection", "Staff cannot be assigned duties on days they are marked out of school."),
    ("Respect Protected Periods", "Teachers must keep their personal protected-period allowance from Teaching Loads where possible."),
    ("Max Duties Per Week", "No staff member should exceed the configured maximum duties per week where possible."),
    ("No Consecutive Heavy Duties", "Avoid assigning Isolation, Lunch, and Detention heavy duties to the same person repeatedly or back-to-back where possible."),
    ("Trained Staff Only for Isolation", "Isolation should only be assigned to staff marked/understood as suitable for isolation duty."),
    ("Respect Prepopulation", "Auto-build must not overwrite manually pre-populated assignments."),
    ("Isolation = SLT or Pastoral Only", "Isolation duties can only be assigned to SLT or Pastoral staff."),
    ("Period 1 Isolation = SLT Only", "Period 1 Isolation is covered by SLT, not Pastoral. Only selected/participating SLT should be used."),
    ("First Duty Other Periods = SLT or Pastoral", "First Duty outside Tutor Time can be assigned to any available SLT or Pastoral member."),
    ("Late Detention at Break = Pastoral", "Late Detention at Break is staffed by a Pastoral department member and should be distributed evenly if multiple Pastoral staff are available."),
    ("Period 4 Mutual Exclusion", "Staff assigned to Period 4 Lunch cannot also be assigned to 4A, 4B, or 4C duties on the same day, and vice versa."),
    ("Lunch Fill Order", "After suitable teachers are used for lunch duty, fill remaining lunch spaces from ESLT, then Chaplaincy, then Admin, and finally SLT."),
    ("Lunch Off-Duty / Rest Protection", "During Period 4 split lunch, Pastoral allocation must preserve the intended rest/off-duty rotation across 4A, 4B, and 4C where possible."),
    ("Room 90 Optional", "Room 90 is the lowest-priority duty. Auto-assign may leave it blank and should not log it as an insufficient-staff problem."),
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
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS custom_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            active INTEGER DEFAULT 1,
            duty_scope TEXT DEFAULT 'Any',
            staff_scope TEXT DEFAULT 'Any',
            condition_type TEXT NOT NULL,
            condition_value TEXT,
            priority TEXT DEFAULT 'Hard',
            notes TEXT,
            is_archived INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            last_updated TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            reason TEXT,
            snapshot_json TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )


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
    for name, description in DEFAULT_RULES:
        conn.execute("UPDATE rules SET description = COALESCE(NULLIF(description, ''), ?) WHERE name = ?", (description, name))
