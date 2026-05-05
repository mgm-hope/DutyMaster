from __future__ import annotations

import re
import sqlite3
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .constants import (
    DUTY_LABELS,
    DUTY_ORDER,
    DUTY_SECTIONS,
    EXPORTS_DIR,
    MASTER_TEMPLATE_PATH,
    ROTA_DAYS,
    ROTA_WEEKS,
)


def parse_timetable(xlsx_path: Path) -> dict:
    wb = load_workbook(xlsx_path, data_only=False)
    ws = wb.active
    school_name = ws.cell(row=3, column=2).value or "Unknown School"

    week1_label = None
    week2_label = None
    for row in range(1, 50):
        for col in range(1, 50):
            val = ws.cell(row=row, column=col).value
            if val and "Week 1" in str(val):
                week1_label = str(val)
            if val and "Week 2" in str(val):
                week2_label = str(val)

    day_map = {
        "Monday": "Mon",
        "Tuesday": "Tue",
        "Wednesday": "Wed",
        "Thursday": "Thu",
        "Friday": "Fri",
    }
    teaching_periods = {"Tutor", "1", "2", "3", "4", "5", "6"}
    all_period_headers = []
    current_day = None
    current_week = 1

    for col in range(1, ws.max_column + 1):
        day_val = ws.cell(row=13, column=col).value
        period_val = ws.cell(row=14, column=col).value
        if day_val in day_map:
            if day_val == "Monday" and current_day == "Fri":
                current_week = 2
            current_day = day_map[day_val]
        if current_day and period_val:
            all_period_headers.append(
                {"week": current_week, "day": current_day, "period": str(period_val), "col": col}
            )

    period_columns = []
    for idx, header in enumerate(all_period_headers):
        if header["period"] in teaching_periods:
            next_col = all_period_headers[idx + 1]["col"] if idx + 1 < len(all_period_headers) else ws.max_column + 1
            period_columns.append({**header, "end_col": next_col - 1})

    teacher_period_rows = []
    real_teacher_day_presence: dict[str, set[tuple[int, str]]] = {}
    seen_period_rows = set()
    for period_info in period_columns:
        for row in range(15, ws.max_row + 1):
            for col in range(period_info["col"], period_info["end_col"] + 1):
                val = ws.cell(row=row, column=col).value
                if val and isinstance(val, str):
                    initials = val.strip()
                    if re.match(r"^[A-Z]{3}$", initials):
                        key = (initials, period_info["week"], period_info["day"], period_info["period"])
                        if key not in seen_period_rows:
                            seen_period_rows.add(key)
                            real_teacher_day_presence.setdefault(initials, set()).add((period_info["week"], period_info["day"]))
                            teacher_period_rows.append(
                                {
                                    "teacher_initials": initials,
                                    "week": period_info["week"],
                                    "day": period_info["day"],
                                    "period": period_info["period"],
                                    "source_row": row,
                                    "source_col": col,
                                }
                            )

    teachers_seen = {row["teacher_initials"] for row in teacher_period_rows}
    occupied_slots = {
        (row["teacher_initials"], row["week"], row["day"], row["period"])
        for row in teacher_period_rows
    }
    for initials in sorted(teachers_seen):
        for week in [1, 2]:
            if (week, "Mon") in real_teacher_day_presence.get(initials, set()) and (initials, week, "Mon", "6") not in occupied_slots:
                teacher_period_rows.append(
                    {
                        "teacher_initials": initials,
                        "week": week,
                        "day": "Mon",
                        "period": "6",
                        "source_row": -1,
                        "source_col": -1,
                    }
                )

    teacher_counts: dict[str, float] = {}
    for row in teacher_period_rows:
        weight = 0.5 if row["period"] == "Tutor" else 1
        teacher_counts[row["teacher_initials"]] = teacher_counts.get(row["teacher_initials"], 0) + weight

    day_order = [(1, "Mon"), (1, "Tue"), (1, "Wed"), (1, "Thu"), (1, "Fri"),
                 (2, "Mon"), (2, "Tue"), (2, "Wed"), (2, "Thu"), (2, "Fri")]
    teachers_data = []
    for initials in sorted(teacher_counts):
        total = teacher_counts[initials]
        days_in_school = "".join(
            "1" if day_key in real_teacher_day_presence.get(initials, set()) else "0"
            for day_key in day_order
        )
        days_out = days_in_school.count("0")
        max_load = 70 - (6.5 * days_out)
        teachers_data.append(
            {
                "initials": initials,
                "full_name": f"Teacher {initials}",
                "is_teaching": 1 if total > 0 else 0,
                "lessons_week1": 0,
                "lessons_week2": total,
                "total_lessons": total,
                "non_contact": max(0, max_load - total),
                "protected_periods": 6,
                "classification": "Teacher",
                "is_part_time": 1 if "0" in days_in_school else 0,
                "days_in_school": days_in_school,
            }
        )

    return {
        "school_name": school_name,
        "week1_label": week1_label or "Week 1",
        "week2_label": week2_label or "Week 2 (inferred)",
        "teachers": teachers_data,
        "teacher_periods": teacher_period_rows,
    }


def reset_upload_data(conn: sqlite3.Connection) -> None:
    for table in ["timetable_meta", "teachers", "teacher_periods", "rota_assignments"]:
        conn.execute(f"DELETE FROM {table}")
    conn.commit()


def save_parsed_timetable(conn: sqlite3.Connection, parsed: dict) -> None:
    reset_upload_data(conn)
    conn.execute(
        "INSERT INTO timetable_meta (school_name, week1_label, week2_label, uploaded_at) VALUES (?, ?, ?, ?)",
        (parsed["school_name"], parsed["week1_label"], parsed["week2_label"], datetime.now().isoformat()),
    )
    for t in parsed["teachers"]:
        conn.execute(
            """
            INSERT INTO teachers (
                initials, full_name, is_teaching, lessons_week1, lessons_week2, total_lessons,
                non_contact, protected_periods, classification, is_part_time, days_in_school, last_updated
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                t["initials"], t["full_name"], t["is_teaching"], t["lessons_week1"], t["lessons_week2"],
                t["total_lessons"], t["non_contact"], t["protected_periods"], t["classification"],
                t["is_part_time"], t["days_in_school"], datetime.now().isoformat(),
            ),
        )
    for row in parsed["teacher_periods"]:
        conn.execute(
            """
            INSERT OR IGNORE INTO teacher_periods
            (teacher_initials, week, day, period, source_row, source_col)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (row["teacher_initials"], row["week"], row["day"], row["period"], row["source_row"], row["source_col"]),
        )
    ensure_duty_event_rows(conn)
    conn.commit()


def ensure_duty_event_rows(conn: sqlite3.Connection) -> None:
    for week in ROTA_WEEKS:
        for day in ROTA_DAYS:
            for code in DUTY_LABELS:
                conn.execute(
                    "INSERT OR IGNORE INTO rota_assignments (week, day, period) VALUES (?, ?, ?)",
                    (week, day, code),
                )
    conn.commit()


def event_to_timetable_period(code: str) -> str | None:
    if code.startswith("Tutor_"):
        return "Tutor"
    for number in ["1", "2", "3", "5", "6"]:
        if code.startswith(f"P{number}_"):
            return number
    if code.startswith("P4"):
        return "4"
    return None


def duty_time_group(code: str) -> str:
    if code == "Gate":
        return "Gate"
    for prefix, group in [
        ("Tutor_", "Tutor"), ("P1_", "P1"), ("P2_", "P2"), ("Break_", "Break"),
        ("P3_", "P3"), ("P4A_", "P4A"), ("P4B_", "P4B"), ("P4C_", "P4C"),
        ("P4_", "P4_Lunch"), ("P5_", "P5"), ("P6_", "P6"), ("P7_", "P7"),
    ]:
        if code.startswith(prefix):
            return group
    return code


def groups_conflict(existing_group: str, requested_group: str) -> bool:
    if existing_group == requested_group:
        return True
    p4_phases = {"P4A", "P4B", "P4C"}
    return (
        existing_group == "P4_Lunch" and requested_group in p4_phases
    ) or (
        requested_group == "P4_Lunch" and existing_group in p4_phases
    )


def same_time_assignment_exists(conn: sqlite3.Connection, initials: str, week: int, day: str, code: str) -> bool:
    current_group = duty_time_group(code)
    rows = conn.execute(
        """
        SELECT period FROM rota_assignments
        WHERE staff_initials = ? AND week = ? AND day = ?
        """,
        (initials, week, day),
    ).fetchall()
    return any(groups_conflict(duty_time_group(row["period"]), current_group) for row in rows)


def teacher_available(conn: sqlite3.Connection, initials: str, week: int, day: str, code: str) -> bool:
    row = conn.execute("SELECT days_in_school FROM teachers WHERE initials = ?", (initials,)).fetchone()
    if not row:
        return False
    days = (row["days_in_school"] or "1111111111").ljust(10, "1")[:10]
    idx = (week - 1) * 5 + ROTA_DAYS.index(day)
    if days[idx] != "1":
        return False
    teaching_period = event_to_timetable_period(code)
    if teaching_period:
        busy = conn.execute(
            """
            SELECT 1 FROM teacher_periods
            WHERE teacher_initials = ? AND week = ? AND day = ? AND period = ?
            LIMIT 1
            """,
            (initials, week, day, teaching_period),
        ).fetchone()
        if busy:
            return False
    return not same_time_assignment_exists(conn, initials, week, day, code)


def additional_available(conn: sqlite3.Connection, initials: str, week: int, day: str, code: str) -> bool:
    row = conn.execute(
        """
        SELECT status, COALESCE(days_in_school, '1111111111') AS days_in_school
        FROM additional_staff
        WHERE initials = ? AND COALESCE(is_archived, 0) = 0
        """,
        (initials,),
    ).fetchone()
    if not row or (row["status"] or "Active") != "Active":
        return False
    days = (row["days_in_school"] or "1111111111").ljust(10, "1")[:10]
    idx = (week - 1) * 5 + ROTA_DAYS.index(day)
    if days[idx] != "1":
        return False
    return not same_time_assignment_exists(conn, initials, week, day, code)


def role_priority(role: str) -> int:
    return {"Pastoral": 0, "SLT": 1, "HOF": 2, "Teacher": 3, "ESLT": 4, "Chaplaincy": 5, "Admin": 6}.get(role, 9)


def duty_is_heavy(code: str) -> bool:
    return "Isolation" in code or "Lunch" in code or "Detention" in code


def duty_is_optional(code: str) -> bool:
    return "Room_90" in code


def duty_scope_matches(code: str, scope: str) -> bool:
    scope = scope or "Any"
    if scope == "Any":
        return True
    if scope == code:
        return True
    scope_prefixes = {
        "Gate Duty": ["Gate"],
        "Tutor Time": ["Tutor_"],
        "Period 1": ["P1_"],
        "Period 2": ["P2_"],
        "Break": ["Break_"],
        "Period 3": ["P3_"],
        "Period 4 Lunch": ["P4_Lunch_"],
        "Period 4A": ["P4A_"],
        "Period 4B": ["P4B_"],
        "Period 4C": ["P4C_"],
        "Period 5": ["P5_"],
        "Period 6": ["P6_"],
        "Period 7": ["P7_"],
        "Isolation Duties": ["Tutor_Isolation", "P1_Isolation", "P2_Isolation", "Break_Isolation", "P3_Isolation", "P4A_Isolation", "P4B_Isolation", "P4C_Isolation", "P5_Isolation", "P6_Isolation", "P7_Isolation"],
        "Lunch and Detention": ["P4_Lunch_", "Break_Late_Detention", "P7_Detention"],
        "Heavy Duties": ["Isolation", "Lunch", "Detention"],
    }
    prefixes = scope_prefixes.get(scope, [])
    return any(code.startswith(prefix) or prefix in code for prefix in prefixes)


def staff_scope_matches(initials: str, role: str, scope: str) -> bool:
    scope = scope or "Any"
    if scope == "Any":
        return True
    if scope.startswith("Staff:"):
        return initials == scope.split(":", 1)[1]
    return role == scope


def staff_week_duty_count(conn: sqlite3.Connection, initials: str, week: int) -> int:
    row = conn.execute(
        "SELECT COUNT(*) AS count FROM rota_assignments WHERE staff_initials = ? AND week = ?",
        (initials, week),
    ).fetchone()
    return int(row["count"] or 0)


def has_heavy_duty_same_day(conn: sqlite3.Connection, initials: str, week: int, day: str) -> bool:
    rows = conn.execute(
        "SELECT period FROM rota_assignments WHERE staff_initials = ? AND week = ? AND day = ?",
        (initials, week, day),
    ).fetchall()
    return any(duty_is_heavy(row["period"]) for row in rows)


def teacher_available_periods(conn: sqlite3.Connection, initials: str) -> float:
    row = conn.execute(
        """
        SELECT total_lessons, protected_periods, days_in_school
        FROM teachers
        WHERE initials = ?
        """,
        (initials,),
    ).fetchone()
    if not row:
        return 999
    days = (row["days_in_school"] or "1111111111").ljust(10, "1")[:10]
    max_load = 70 - (6.5 * days.count("0"))
    current_duties = conn.execute(
        "SELECT COUNT(*) AS count FROM rota_assignments WHERE staff_initials = ?",
        (initials,),
    ).fetchone()["count"]
    return max_load - float(row["total_lessons"] or 0) - int(row["protected_periods"] or 0) - int(current_duties or 0)


def custom_rules_allow(conn: sqlite3.Connection, initials: str, role: str, week: int, day: str, code: str) -> bool:
    rules = conn.execute(
        """
        SELECT duty_scope, staff_scope, condition_type, condition_value, priority
        FROM custom_rules
        WHERE active = 1 AND COALESCE(is_archived, 0) = 0
        """
    ).fetchall()
    for rule in rules:
        if (rule["priority"] or "Hard") != "Hard":
            continue
        if not duty_scope_matches(code, rule["duty_scope"]):
            continue
        if not staff_scope_matches(initials, role, rule["staff_scope"]):
            continue
        condition = rule["condition_type"]
        value = (rule["condition_value"] or "").strip()
        if condition == "exclude":
            return False
        if condition == "min_available_periods":
            try:
                minimum = float(value)
            except ValueError:
                minimum = 0
            if teacher_available_periods(conn, initials) < minimum:
                return False
        elif condition == "max_duties_per_week":
            try:
                maximum = int(float(value))
            except ValueError:
                maximum = 0
            if staff_week_duty_count(conn, initials, week) >= maximum:
                return False
        elif condition == "no_heavy_same_day" and duty_is_heavy(code):
            if has_heavy_duty_same_day(conn, initials, week, day):
                return False
    return True


def role_priority_for_duty(code: str) -> list[str]:
    if code == "Gate":
        return ["SLT"]
    if code in {"Tutor_1st_Duty", "Tutor_AOW", "P1_Isolation"}:
        return ["SLT"]
    if code == "Break_Duty_Lead":
        return ["SLT", "Pastoral"]
    if code == "Break_Late_Detention":
        return ["Pastoral"]
    if "Pastoral_Support" in code or "Room_90" in code:
        return ["Pastoral", "SLT"]
    if "Isolation" in code:
        return ["Pastoral", "SLT"]
    if "First_Duty" in code:
        return ["Pastoral", "SLT"]
    if code == "P4_Lunch_1":
        return ["Pastoral"]
    if code.startswith("P4_Lunch_"):
        return ["Teacher", "HOF", "ESLT", "Chaplaincy", "Admin", "SLT"]
    if code.startswith("P7_Detention"):
        return ["Pastoral", "SLT", "Teacher", "HOF"]
    return ["Pastoral", "SLT", "Teacher", "HOF", "ESLT", "Chaplaincy", "Admin"]


def strict_assignment_allowed(
    conn: sqlite3.Connection,
    initials: str,
    role: str,
    week: int,
    day: str,
    code: str,
    source: str | None = None,
) -> bool:
    allowed_roles = role_priority_for_duty(code)
    if role not in allowed_roles:
        return False
    if source is None:
        source = "teacher" if conn.execute("SELECT 1 FROM teachers WHERE initials = ?", (initials,)).fetchone() else "additional"
    if source == "teacher":
        if not teacher_available(conn, initials, week, day, code):
            return False
    else:
        if not additional_available(conn, initials, week, day, code):
            return False
    return custom_rules_allow(conn, initials, role, week, day, code)


def available_staff(conn: sqlite3.Connection, week: int, day: str, code: str) -> list[dict]:
    staff = []
    for row in conn.execute("SELECT initials, full_name, classification FROM teachers ORDER BY initials").fetchall():
        if strict_assignment_allowed(conn, row["initials"], row["classification"], week, day, code, "teacher"):
            staff.append({"initials": row["initials"], "name": row["full_name"], "role": row["classification"]})
    for row in conn.execute(
        """
        SELECT initials, full_name, category FROM additional_staff
        WHERE COALESCE(is_archived, 0) = 0 AND COALESCE(status, 'Active') = 'Active'
        ORDER BY category, initials
        """
    ).fetchall():
        if strict_assignment_allowed(conn, row["initials"], row["category"], week, day, code, "additional"):
            staff.append({"initials": row["initials"], "name": row["full_name"], "role": row["category"]})
    return sorted(staff, key=lambda item: (role_priority(item["role"]), item["initials"]))


def clear_conflicting_p4_assignments(conn: sqlite3.Connection, initials: str, week: int, day: str, code: str) -> int:
    if code.startswith("P4_Lunch_"):
        cursor = conn.execute(
            """
            UPDATE rota_assignments SET staff_initials = NULL, staff_type = NULL, last_updated = ?
            WHERE week = ? AND day = ? AND staff_initials = ?
              AND (period LIKE 'P4A_%' OR period LIKE 'P4B_%' OR period LIKE 'P4C_%')
            """,
            (datetime.now().isoformat(), week, day, initials),
        )
        return cursor.rowcount
    if code.startswith("P4A_") or code.startswith("P4B_") or code.startswith("P4C_"):
        cursor = conn.execute(
            """
            UPDATE rota_assignments SET staff_initials = NULL, staff_type = NULL, last_updated = ?
            WHERE week = ? AND day = ? AND staff_initials = ? AND period LIKE 'P4_Lunch_%'
            """,
            (datetime.now().isoformat(), week, day, initials),
        )
        return cursor.rowcount
    return 0


def assign_staff(conn: sqlite3.Connection, week: int, day: str, code: str, initials: str, role: str) -> int:
    if not strict_assignment_allowed(conn, initials, role, week, day, code):
        return -1
    cleared = 0
    conn.execute(
        """
        UPDATE rota_assignments
        SET staff_initials = ?, staff_type = ?, last_updated = ?
        WHERE week = ? AND day = ? AND period = ?
        """,
        (initials, role, datetime.now().isoformat(), week, day, code),
    )
    conn.commit()
    return cleared


def get_p4_lunch_conflicts(conn: sqlite3.Connection) -> list[sqlite3.Row]:
    return conn.execute(
        """
        SELECT week, day, staff_initials, GROUP_CONCAT(period, ', ') AS duties
        FROM rota_assignments
        WHERE staff_initials IS NOT NULL
          AND (period LIKE 'P4_Lunch_%' OR period LIKE 'P4A_%' OR period LIKE 'P4B_%' OR period LIKE 'P4C_%')
        GROUP BY week, day, staff_initials
        HAVING SUM(CASE WHEN period LIKE 'P4_Lunch_%' THEN 1 ELSE 0 END) > 0
           AND SUM(CASE WHEN period LIKE 'P4A_%' OR period LIKE 'P4B_%' OR period LIKE 'P4C_%' THEN 1 ELSE 0 END) > 0
        ORDER BY week, day, staff_initials
        """
    ).fetchall()


def repair_p4_lunch_conflicts(conn: sqlite3.Connection) -> int:
    conflicts = get_p4_lunch_conflicts(conn)
    cleared = 0
    for row in conflicts:
        cursor = conn.execute(
            """
            UPDATE rota_assignments
            SET staff_initials = NULL, staff_type = NULL, last_updated = ?
            WHERE week = ? AND day = ? AND staff_initials = ?
              AND (period LIKE 'P4A_%' OR period LIKE 'P4B_%' OR period LIKE 'P4C_%')
            """,
            (datetime.now().isoformat(), row["week"], row["day"], row["staff_initials"]),
        )
        cleared += cursor.rowcount
    conn.commit()
    return cleared


def duty_sort_key(slot: sqlite3.Row | tuple[int, str, str]) -> tuple:
    week, day, code = (slot["week"], slot["day"], slot["period"]) if hasattr(slot, "keys") else slot
    priority = 50
    if code == "Gate":
        priority = 0
    elif code == "P1_Isolation":
        priority = 1
    elif code == "Tutor_1st_Duty":
        priority = 2
    elif code == "Break_Duty_Lead":
        priority = 3
    elif code.startswith("P7_Detention"):
        priority = 4
    elif code.startswith("P4_Lunch"):
        priority = 5
    elif "Isolation" in code:
        priority = 6
    elif duty_is_optional(code):
        priority = 99
    return (priority, week, ROTA_DAYS.index(day), DUTY_ORDER.get(code, 999))


def auto_assign_empty_slots(conn: sqlite3.Connection) -> dict:
    ensure_duty_event_rows(conn)
    pre_repaired = repair_p4_lunch_conflicts(conn)

    candidates = []
    for row in conn.execute(
        """
        SELECT initials, classification AS role, total_lessons, protected_periods, days_in_school
        FROM teachers
        """
    ).fetchall():
        days = (row["days_in_school"] or "1111111111").ljust(10, "1")[:10]
        max_load = 70 - (6.5 * days.count("0"))
        current_duties = conn.execute("SELECT COUNT(*) AS count FROM rota_assignments WHERE staff_initials = ?", (row["initials"],)).fetchone()["count"]
        heavy_rows = conn.execute("SELECT period FROM rota_assignments WHERE staff_initials = ?", (row["initials"],)).fetchall()
        heavy_count = sum(1 for heavy in heavy_rows if duty_is_heavy(heavy["period"]))
        candidates.append(
            {
                "initials": row["initials"],
                "role": row["role"],
                "available": max_load - float(row["total_lessons"] or 0) - int(row["protected_periods"] or 0) - current_duties,
                "duties": current_duties,
                "heavy": heavy_count,
                "source": "teacher",
            }
        )

    for row in conn.execute(
        """
        SELECT initials, category AS role FROM additional_staff
        WHERE COALESCE(is_archived, 0) = 0 AND COALESCE(status, 'Active') = 'Active'
        """
    ).fetchall():
        current_duties = conn.execute("SELECT COUNT(*) AS count FROM rota_assignments WHERE staff_initials = ?", (row["initials"],)).fetchone()["count"]
        heavy_rows = conn.execute("SELECT period FROM rota_assignments WHERE staff_initials = ?", (row["initials"],)).fetchall()
        heavy_count = sum(1 for heavy in heavy_rows if duty_is_heavy(heavy["period"]))
        candidates.append(
            {
                "initials": row["initials"],
                "role": row["role"],
                "available": 999,
                "duties": current_duties,
                "heavy": heavy_count,
                "source": "additional",
            }
        )

    empty_slots = conn.execute(
        "SELECT week, day, period FROM rota_assignments WHERE staff_initials IS NULL"
    ).fetchall()
    assigned = 0
    issues = []
    for slot in sorted(empty_slots, key=duty_sort_key):
        week, day, code = slot["week"], slot["day"], slot["period"]
        allowed_roles = role_priority_for_duty(code)
        eligible = []
        for cand in candidates:
            if not strict_assignment_allowed(conn, cand["initials"], cand["role"], week, day, code, cand["source"]):
                continue
            eligible.append((allowed_roles.index(cand["role"]), -cand["available"], cand["duties"], cand["heavy"], cand["initials"], cand))
        if not eligible:
            if not duty_is_optional(code):
                issues.append(slot)
            continue
        eligible.sort()
        chosen = eligible[0][-1]
        conn.execute(
            """
            UPDATE rota_assignments
            SET staff_initials = ?, staff_type = ?, last_updated = ?
            WHERE week = ? AND day = ? AND period = ?
            """,
            (chosen["initials"], chosen["role"], datetime.now().isoformat(), week, day, code),
        )
        chosen["duties"] += 1
        chosen["available"] -= 1
        if duty_is_heavy(code):
            chosen["heavy"] += 1
        assigned += 1

    for slot in issues:
        conn.execute(
            """
            INSERT INTO problem_log(issue_type, description, week, day, period)
            VALUES (?, ?, ?, ?, ?)
            """,
            ("Insufficient Staff", f"Auto-assign could not fill {DUTY_LABELS.get(slot['period'], slot['period'])}", slot["week"], slot["day"], slot["period"]),
        )
    conn.commit()
    post_repaired = repair_p4_lunch_conflicts(conn)
    return {"assigned": assigned, "issues": len(issues), "repaired": pre_repaired + post_repaired}


def build_master_style_workbook(conn: sqlite3.Connection) -> BytesIO:
    ensure_duty_event_rows(conn)
    assignment_rows = conn.execute("SELECT week, day, period, staff_initials FROM rota_assignments").fetchall()
    assignments = {(row["week"], row["day"], row["period"]): row["staff_initials"] or "" for row in assignment_rows}

    template_path = MASTER_TEMPLATE_PATH if MASTER_TEMPLATE_PATH.exists() else Path("2025 MASTER DUTY FINAL.xlsx")
    if template_path.exists():
        wb = load_workbook(template_path)
        ws = wb["Master"]
        day_columns = {
            (1, "Mon"): 2, (1, "Tue"): 3, (1, "Wed"): 4, (1, "Thu"): 5, (1, "Fri"): 6,
            (2, "Mon"): 7, (2, "Tue"): 8, (2, "Wed"): 9, (2, "Thu"): 10, (2, "Fri"): 11,
        }
        staff_rows = [3, 5, 7, 8, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 34, 35, 37, 39,
                      41, 43, 45, 47, 49, 51, 53, 55, 57, 59, 60, 61, 62, 63, 64, 65, 67, 71,
                      73, 75, 77, 80, 82, 84, 86, 88, 90, 92, 94, 96, 98, 102]
        for row in staff_rows:
            for col in range(2, 12):
                ws.cell(row, col).value = None
        for row in range(12, ws.max_row + 1):
            for col in range(17, 24):
                ws.cell(row, col).value = None
        master_row_map = {
            "Tutor_1st_Duty": 3, "Tutor_AOW": 7, "Tutor_Pastoral_Support": 11, "Tutor_Room_90": 13,
            "P1_First_Duty": 15, "P1_Pastoral_Support": 17, "P1_Room_90": 19, "P1_Isolation": 21,
            "P2_First_Duty": 23, "P2_Pastoral_Support": 25, "P2_Room_90": 27, "P2_Isolation": 29,
            "Break_Duty_Lead": 31, "Break_Late_Detention": 34, "Break_Pastoral_Support": 37, "Break_Room_90": 39,
            "P3_First_Duty": 41, "P3_Pastoral_Support": 43, "P3_Room_90": 45, "P3_Isolation": 47,
            "P4A_First_Duty": 49, "P4A_Pastoral_Support": 51, "P4A_Isolation": 55,
            "P4_Lunch_1": 57, "P4_Lunch_2": 59, "P4_Lunch_3": 60, "P4_Lunch_4": 61,
            "P4_Lunch_5": 62, "P4_Lunch_6": 63, "P4_Lunch_7": 64,
            "P4B_First_Duty": 65, "P4B_Pastoral_Support": 67, "P4B_Isolation": 71,
            "P4C_First_Duty": 73, "P4C_Pastoral_Support": 75, "P4C_Isolation": 77, "P4C_Rest_Break": 80,
            "P5_First_Duty": 82, "P5_Pastoral_Support": 84, "P5_Room_90": 86, "P5_Isolation": 88,
            "P6_First_Duty": 90, "P6_Pastoral_Support": 92, "P6_Room_90": 94, "P6_Isolation": 96,
            "P7_Homework_Club": 98, "P7_Detention_Drop_In": 102,
        }
        for code, row in master_row_map.items():
            for (week, day), col in day_columns.items():
                ws.cell(row, col).value = assignments.get((week, day, code), "")
        if "DutyMaster extra duties" in wb.sheetnames:
            del wb["DutyMaster extra duties"]
        extra = wb.create_sheet("DutyMaster extra duties")
        extra.append(["Week", "Day", "Duty", "Staff"])
        extra_codes = ["Gate", "Tutor_Isolation", "Break_Isolation", "P4A_Rest_Break", "P4B_Rest_Break", "P7_Isolation", "P7_Detention_1", "P7_Detention_2"]
        for code in extra_codes:
            for week in ROTA_WEEKS:
                for day in ROTA_DAYS:
                    extra.append([week, day, DUTY_LABELS.get(code, code), assignments.get((week, day, code), "")])
        for col in range(1, 5):
            extra.column_dimensions[get_column_letter(col)].width = 18
            extra.cell(1, col).font = Font(bold=True)
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    wb = Workbook()
    ws = wb.active
    ws.title = "Master"
    fills = {
        "header": PatternFill("solid", fgColor="1F4E78"),
        "section": PatternFill("solid", fgColor="E2F0D9"),
    }
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(["Duty / Event", "Week 1 Mon", "Week 1 Tue", "Week 1 Wed", "Week 1 Thu", "Week 1 Fri",
               "Week 2 Mon", "Week 2 Tue", "Week 2 Wed", "Week 2 Thu", "Week 2 Fri"])
    for cell in ws[1]:
        cell.fill = fills["header"]
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for section, events in DUTY_SECTIONS:
        ws.append([section])
        ws.cell(ws.max_row, 1).fill = fills["section"]
        for code, label in events:
            row = [label]
            for week in ROTA_WEEKS:
                for day in ROTA_DAYS:
                    row.append(assignments.get((week, day, code), ""))
            ws.append(row)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for idx in range(1, 12):
        ws.column_dimensions[get_column_letter(idx)].width = 16
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
