# DutyMaster Online

This folder is the Railway-ready online companion to the local Streamlit/PC version in the parent `DutyMaster` folder.

The split is deliberate:

- `dutymaster.py` remains the PC/desktop Streamlit version.
- `DutyMaster Online/` is the hosted FastAPI version for Railway.

## What is ready

- FastAPI app with `main.py -> app.web:app`
- Password login using Railway environment variables
- SQLite persistence using a Railway `/data` volume
- Timetable upload using the existing DutyMaster parsing logic
- Staff list upload for initials to full names
- Teaching loads page with protected periods, classification, and days-in-school editing
- Additional staffing page for Pastoral, Admin, ESLT, Chaplaincy, and non-teaching SLT
- Pre-built duty events grid using the DutyMaster staffing structure
- Manual assignment with Period 4 lunch vs 4A/4B/4C clash protection
- Proposed timetable view
- Master Duty Excel download using the uploaded master template where available
- Narrative & Problems log

## Local run

From this folder:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:DUTYMASTER_PASSWORD = "change-this"
$env:DUTYMASTER_SECRET_KEY = "use-a-long-random-string"
$env:DUTYMASTER_DATA_DIR = ".\data"
uvicorn main:app --reload
```

Open:

```text
http://127.0.0.1:8000
```

## Railway environment variables

Set these on the new Railway service:

```text
DUTYMASTER_PASSWORD=your-shared-password
DUTYMASTER_SECRET_KEY=long-random-secret
DUTYMASTER_DATA_DIR=/data
DUTYMASTER_DB_PATH=/data/dutymaster_online.sqlite3
DUTYMASTER_UPLOAD_DIR=/data/uploads
DUTYMASTER_EXPORT_DIR=/data/exports
DUTYMASTER_MASTER_TEMPLATE=/data/2025 MASTER DUTY FINAL.xlsx
```

## Railway volume

Add a persistent Railway volume to this new service and mount it at:

```text
/data
```

The SQLite database, uploads, exports, and master duty template will live there.

## Deployment into `SEND TOOLS`

Use a separate Railway service inside the existing `SEND TOOLS` project:

1. Open Railway and go to the existing `SEND TOOLS` project.
2. Add a **new service** from this folder/repository.
3. Make the service root point at `DutyMaster Online` if the repo contains both PC and online folders.
4. Railway will use `railway.json`:

```text
uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
```

5. Add the environment variables above.
6. Add and mount the `/data` volume.
7. Deploy the service.
8. Open the public Railway URL and log in.
9. Upload:
   - the timetable export
   - the staff list if needed
   - `2025 MASTER DUTY FINAL.xlsx` as the master export template
10. Add the live Railway URL to the existing tools homepage.

## Homepage link

Once live, add a card/link to the tools homepage using the Railway public URL, for example:

```text
DutyMaster Online
School duty rota builder and export tool
https://your-duty-master-service.up.railway.app
```

## Notes

This is the first practical online cut. It is intentionally separate from the PC app so the desktop version stays safe while the Railway version evolves.

