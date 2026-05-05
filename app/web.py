from __future__ import annotations

import os
import secrets
import shutil
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from .constants import (
    APP_DIR,
    DATA_DIR,
    DUTY_LABELS,
    DUTY_ORDER,
    DUTY_SECTIONS,
    EXPORTS_DIR,
    MASTER_TEMPLATE_PATH,
    ROTA_DAYS,
    ROTA_WEEKS,
    UPLOADS_DIR,
)
from .database import DB_PATH, get_connection, initialise_database
from .services import (
    assign_staff,
    auto_assign_empty_slots,
    available_staff,
    build_master_style_workbook,
    ensure_duty_event_rows,
    get_p4_lunch_conflicts,
    parse_timetable,
    reset_upload_data,
    repair_p4_lunch_conflicts,
    save_parsed_timetable,
)


PACKAGE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = PACKAGE_DIR / "templates"
STATIC_DIR = PACKAGE_DIR / "static"


def _ensure_bootstrap() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    conn = initialise_database(DB_PATH)
    ensure_duty_event_rows(conn)
    conn.close()


def _new_connection() -> sqlite3.Connection:
    conn = get_connection(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def _conn_context() -> Iterator[sqlite3.Connection]:
    conn = _new_connection()
    try:
        yield conn
    finally:
        conn.close()


def _password_value() -> str:
    return os.getenv("DUTYMASTER_PASSWORD", "changeme123!")


def _is_logged_in(request: Request) -> bool:
    return bool(request.session.get("logged_in"))


def _require_login(request: Request) -> RedirectResponse | None:
    if _is_logged_in(request):
        return None
    return RedirectResponse("/login", status_code=303)


def _flash(request: Request, message: str) -> None:
    request.session["flash"] = message


def _get_flash(request: Request) -> str:
    return request.session.pop("flash", "")


def _base_context(request: Request, title: str, **extra) -> dict:
    return {
        "request": request,
        "title": title,
        "flash": _get_flash(request),
        "is_logged_in": _is_logged_in(request),
        "days": ROTA_DAYS,
        "weeks": ROTA_WEEKS,
        **extra,
    }


def _save_upload(upload: UploadFile, prefix: str) -> Path:
    safe_name = Path(upload.filename or f"{prefix}.xlsx").name
    target = UPLOADS_DIR / f"{prefix}_{secrets.token_hex(4)}_{safe_name}"
    with target.open("wb") as handle:
        shutil.copyfileobj(upload.file, handle)
    return target


_ensure_bootstrap()

app = FastAPI(title="DutyMaster Online")
app.add_middleware(SessionMiddleware, secret_key=os.getenv("DUTYMASTER_SECRET_KEY", "replace-me-on-railway"))
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def root(request: Request):
    if not _is_logged_in(request):
        return RedirectResponse("/login", status_code=303)
    return RedirectResponse("/dashboard", status_code=303)


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse("login.html", _base_context(request, "Login"))


@app.post("/login")
async def login(request: Request, password: str = Form(...)):
    if secrets.compare_digest(password, _password_value()):
        request.session["logged_in"] = True
        _flash(request, "Signed in.")
        return RedirectResponse("/dashboard", status_code=303)
    _flash(request, "Password not recognised.")
    return RedirectResponse("/login", status_code=303)


@app.post("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        meta = conn.execute("SELECT * FROM timetable_meta ORDER BY id DESC LIMIT 1").fetchone()
        teacher_count = conn.execute("SELECT COUNT(*) AS count FROM teachers").fetchone()["count"]
        additional_count = conn.execute("SELECT COUNT(*) AS count FROM additional_staff WHERE COALESCE(is_archived, 0)=0").fetchone()["count"]
        assigned_count = conn.execute("SELECT COUNT(*) AS count FROM rota_assignments WHERE staff_initials IS NOT NULL").fetchone()["count"]
        conflicts = get_p4_lunch_conflicts(conn)
    return templates.TemplateResponse(
        "dashboard.html",
        _base_context(
            request,
            "Dashboard",
            meta=meta,
            teacher_count=teacher_count,
            additional_count=additional_count,
            assigned_count=assigned_count,
            conflicts=conflicts,
            db_path=str(DB_PATH),
        ),
    )


@app.get("/uploads", response_class=HTMLResponse)
def uploads(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        meta = conn.execute("SELECT * FROM timetable_meta ORDER BY id DESC LIMIT 1").fetchone()
    return templates.TemplateResponse(
        "uploads.html",
        _base_context(request, "Uploads", meta=meta, template_path=str(MASTER_TEMPLATE_PATH)),
    )


@app.post("/uploads/timetable")
async def upload_timetable(request: Request, timetable: UploadFile = File(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    target = _save_upload(timetable, "timetable")
    parsed = parse_timetable(target)
    with _conn_context() as conn:
        save_parsed_timetable(conn, parsed)
    _flash(request, f"Uploaded timetable: {len(parsed['teachers'])} teachers parsed. Teaching Loads is ready to review.")
    return RedirectResponse("/teaching-loads", status_code=303)


@app.post("/uploads/staff-list")
async def upload_staff_list(request: Request, staff_list: UploadFile = File(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    target = _save_upload(staff_list, "staff_list")
    from openpyxl import load_workbook
    import csv

    rows = []
    if target.suffix.lower() == ".csv":
        with target.open(newline="", encoding="utf-8-sig") as handle:
            rows = list(csv.DictReader(handle))
    else:
        wb = load_workbook(target)
        ws = wb.active
        headers = [str(cell.value or "").strip() for cell in ws[1]]
        for values in ws.iter_rows(min_row=2, values_only=True):
            rows.append(dict(zip(headers, values)))
    count = 0
    with _conn_context() as conn:
        for row in rows:
            initials = str(row.get("Initials") or row.get("initials") or "").strip().upper()
            full_name = str(row.get("Full Name") or row.get("full_name") or row.get("Name") or "").strip()
            if initials and full_name:
                conn.execute(
                    "INSERT OR REPLACE INTO staff_names(initials, full_name, last_updated) VALUES (?, ?, ?)",
                    (initials, full_name, datetime.now().isoformat()),
                )
                conn.execute("UPDATE teachers SET full_name = ? WHERE initials = ?", (full_name, initials))
                count += 1
        conn.commit()
    _flash(request, f"Imported {count} staff names.")
    return RedirectResponse("/uploads", status_code=303)


@app.post("/uploads/master-template")
async def upload_master_template(request: Request, master_template: UploadFile = File(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    MASTER_TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with MASTER_TEMPLATE_PATH.open("wb") as handle:
        shutil.copyfileobj(master_template.file, handle)
    _flash(request, "Master duty template uploaded for exports.")
    return RedirectResponse("/uploads", status_code=303)


@app.post("/uploads/reset")
def reset_upload(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        reset_upload_data(conn)
    _flash(request, "Uploaded timetable data reset. Additional staffing was kept.")
    return RedirectResponse("/uploads", status_code=303)


@app.get("/teaching-loads", response_class=HTMLResponse)
def teaching_loads(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    try:
        with _conn_context() as conn:
            teachers = conn.execute(
                """
                SELECT t.initials,
                       COALESCE(s.full_name, t.full_name, t.initials) AS full_name,
                       COALESCE(t.total_lessons, 0) AS total_lessons,
                       COALESCE(t.protected_periods, 6) AS protected_periods,
                       COALESCE(t.classification, 'Teacher') AS classification,
                       COALESCE(t.is_part_time, 0) AS is_part_time,
                       COALESCE(t.days_in_school, '1111111111') AS days_in_school
                FROM teachers t
                LEFT JOIN staff_names s ON t.initials = s.initials
                ORDER BY t.initials
                """
            ).fetchall()
        return templates.TemplateResponse("teaching_loads.html", _base_context(request, "Teaching Loads", teachers=teachers))
    except Exception as exc:
        return templates.TemplateResponse(
            "error.html",
            _base_context(
                request,
                "Teaching Loads Error",
                heading="Teaching Loads could not open",
                message=str(exc),
                next_step="Try re-uploading the timetable. If this persists, copy this message from the Railway app.",
            ),
            status_code=500,
        )


@app.post("/teaching-loads/update")
async def update_teacher(
    request: Request,
    initials: str = Form(...),
    protected_periods: int = Form(...),
    classification: str = Form(...),
    days_in_school: str = Form(...),
):
    redirect = _require_login(request)
    if redirect:
        return redirect
    days = "".join("1" if char == "1" else "0" for char in days_in_school)[:10].ljust(10, "1")
    with _conn_context() as conn:
        total = conn.execute("SELECT total_lessons FROM teachers WHERE initials = ?", (initials,)).fetchone()["total_lessons"]
        days_out = days.count("0")
        max_load = 70 - (6.5 * days_out)
        conn.execute(
            """
            UPDATE teachers
            SET protected_periods = ?, classification = ?, days_in_school = ?,
                is_part_time = ?, non_contact = ?, last_updated = ?
            WHERE initials = ?
            """,
            (protected_periods, classification, days, 1 if "0" in days else 0, max(0, max_load - float(total or 0)), datetime.now().isoformat(), initials),
        )
        conn.commit()
    _flash(request, f"Updated {initials}.")
    return RedirectResponse("/teaching-loads", status_code=303)


@app.get("/additional-staff", response_class=HTMLResponse)
def additional_staff(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        staff = conn.execute("SELECT * FROM additional_staff ORDER BY is_archived, category, initials").fetchall()
    return templates.TemplateResponse("additional_staff.html", _base_context(request, "Additional Staffing", staff=staff))


@app.post("/additional-staff/add")
async def add_additional_staff(
    request: Request,
    initials: str = Form(...),
    full_name: str = Form(...),
    category: str = Form(...),
):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        try:
            conn.execute(
                """
                INSERT INTO additional_staff(category, initials, full_name, is_full_time, availability, last_updated)
                VALUES (?, ?, ?, 1, NULL, ?)
                """,
                (category, initials.strip().upper(), full_name.strip(), datetime.now().isoformat()),
            )
            conn.commit()
            _flash(request, f"Added {initials.upper()}.")
        except sqlite3.IntegrityError:
            _flash(request, f"{initials.upper()} already exists.")
    return RedirectResponse("/additional-staff", status_code=303)


@app.post("/additional-staff/status")
async def update_additional_status(request: Request, staff_id: int = Form(...), status: str = Form(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        conn.execute("UPDATE additional_staff SET status = ?, last_updated = ? WHERE id = ?", (status, datetime.now().isoformat(), staff_id))
        conn.commit()
    return RedirectResponse("/additional-staff", status_code=303)


@app.post("/additional-staff/archive")
async def archive_additional(request: Request, staff_id: int = Form(...), archive: int = Form(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        conn.execute("UPDATE additional_staff SET is_archived = ?, last_updated = ? WHERE id = ?", (archive, datetime.now().isoformat(), staff_id))
        conn.commit()
    return RedirectResponse("/additional-staff", status_code=303)


@app.get("/prebuilt", response_class=HTMLResponse)
def prebuilt(request: Request, week: int = 1, day: str = "Mon", period: str | None = None):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        ensure_duty_event_rows(conn)
        assignments = {
            f"{row['week']}:{row['day']}:{row['period']}": row["staff_initials"] or ""
            for row in conn.execute("SELECT week, day, period, staff_initials FROM rota_assignments").fetchall()
        }
        candidates = available_staff(conn, week, day, period) if period else []
        conflicts = get_p4_lunch_conflicts(conn)
    return templates.TemplateResponse(
        "prebuilt.html",
        _base_context(
            request,
            "Pre-Built Duty Events",
            duty_sections=DUTY_SECTIONS,
            duty_labels=DUTY_LABELS,
            assignments=assignments,
            selected={"week": week, "day": day, "period": period},
            candidates=candidates,
            conflicts=conflicts,
        ),
    )


@app.post("/prebuilt/assign")
async def prebuilt_assign(
    request: Request,
    week: int = Form(...),
    day: str = Form(...),
    period: str = Form(...),
    staff_value: str = Form(...),
    scope: str = Form("single"),
):
    redirect = _require_login(request)
    if redirect:
        return redirect
    initials, role = staff_value.split("|", 1)
    targets = [(week, day)]
    if scope == "week":
        targets = [(week, target_day) for target_day in ROTA_DAYS]
    elif scope == "both":
        targets = [(target_week, target_day) for target_week in ROTA_WEEKS for target_day in ROTA_DAYS]
    assigned = 0
    skipped = 0
    cleared = 0
    with _conn_context() as conn:
        for target_week, target_day in targets:
            possible = available_staff(conn, target_week, target_day, period)
            if not any(item["initials"] == initials and item["role"] == role for item in possible):
                skipped += 1
                continue
            cleared += assign_staff(conn, target_week, target_day, period, initials, role)
            assigned += 1
    clear_note = f" Cleared {cleared} Period 4 clash(es)." if cleared else ""
    skip_note = f" Skipped {skipped} unavailable/busy session(s)." if skipped else ""
    _flash(request, f"Assigned {initials} to {assigned} session(s).{skip_note}{clear_note}")
    return RedirectResponse(f"/prebuilt?week={week}&day={day}&period={period}", status_code=303)


@app.post("/prebuilt/clear")
async def prebuilt_clear(request: Request, week: int = Form(...), day: str = Form(...), period: str = Form(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        conn.execute(
            "UPDATE rota_assignments SET staff_initials = NULL, staff_type = NULL, last_updated = ? WHERE week = ? AND day = ? AND period = ?",
            (datetime.now().isoformat(), week, day, period),
        )
        conn.commit()
    return RedirectResponse(f"/prebuilt?week={week}&day={day}&period={period}", status_code=303)


@app.post("/prebuilt/repair-p4")
async def prebuilt_repair_p4(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        cleared = repair_p4_lunch_conflicts(conn)
    _flash(request, f"Repaired Period 4 lunch conflicts. Cleared {cleared} 4A/4B/4C assignment(s).")
    return RedirectResponse("/prebuilt", status_code=303)


@app.post("/prebuilt/auto-assign")
async def prebuilt_auto_assign(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        result = auto_assign_empty_slots(conn)
    _flash(
        request,
        f"Auto-assigned {result['assigned']} slot(s). {result['issues']} need manual review. Repaired {result['repaired']} Period 4 conflict(s).",
    )
    return RedirectResponse("/prebuilt", status_code=303)


@app.get("/proposed", response_class=HTMLResponse)
def proposed(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        assignments = conn.execute(
            """
            SELECT week, day, period, staff_initials, staff_type FROM rota_assignments
            ORDER BY week,
                     CASE day WHEN 'Mon' THEN 1 WHEN 'Tue' THEN 2 WHEN 'Wed' THEN 3 WHEN 'Thu' THEN 4 ELSE 5 END
            """
        ).fetchall()
    return templates.TemplateResponse("proposed.html", _base_context(request, "Proposed Timetable", assignments=assignments, duty_labels=DUTY_LABELS, duty_order=DUTY_ORDER))


@app.get("/proposed/download")
def proposed_download(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        data = build_master_style_workbook(conn)
    return Response(
        content=data.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Proposed_Timetable_Master_Duty_Format.xlsx"},
    )


@app.get("/problems", response_class=HTMLResponse)
def problems(request: Request):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        rows = conn.execute("SELECT * FROM problem_log ORDER BY timestamp DESC").fetchall()
    return templates.TemplateResponse("problems.html", _base_context(request, "Narrative & Problems", rows=rows))


@app.post("/problems/add")
async def problems_add(request: Request, issue_type: str = Form(...), description: str = Form(...)):
    redirect = _require_login(request)
    if redirect:
        return redirect
    with _conn_context() as conn:
        conn.execute(
            "INSERT INTO problem_log(issue_type, description) VALUES (?, ?)",
            (issue_type, description.strip()),
        )
        conn.commit()
    return RedirectResponse("/problems", status_code=303)
