# PostgreSQL Integration Guide

This backend stores report metadata in PostgreSQL while keeping uploaded files,
processed DataFrames, previews, and generated Word/Excel files on the filesystem.
The existing API response shape and report generation flow are preserved.

## What PostgreSQL Stores

- Report/session metadata and status in `reports`
- Upload file metadata and paths in `report_uploads`
- Manual input fields and the full manual payload in `report_manual_inputs`
- Preview file paths in `report_previews`
- Final output file paths in `report_outputs`

The database does not store uploaded file bytes or generated document bytes.

## Start PostgreSQL With Docker Compose

From the repository root:

```bash
docker compose up -d db
```

If your Docker install does not include the Compose plugin, use the equivalent
single-container command:

```bash
docker run --name report-app-postgres \
  -e POSTGRES_DB=report_app_db \
  -e POSTGRES_USER=report_app_user \
  -e POSTGRES_PASSWORD=report_app_password \
  -p 5432:5432 \
  -d postgres:16
```

The local database connection string is:

```text
postgresql+psycopg://report_app_user:report_app_password@localhost:5432/report_app_db
```

If the backend runs inside Docker Compose, use `db` as the host instead of
`localhost`.

## Configure Backend Environment

From `backend/`:

```bash
cp .env.example .env
```

Ensure `.env` contains:

```env
DATABASE_URL=postgresql+psycopg://report_app_user:report_app_password@localhost:5432/report_app_db
APP_ENV=development
REPORT_STORAGE_ROOT=./app/storage
ADMIN_PASSWORD=replace-with-a-strong-admin-password
```

## Install Dependencies

From `backend/`:

```bash
.venv/bin/pip install -r requirements.txt
```

Or, if you are not using the existing virtualenv:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## Run Migrations

From `backend/`:

```bash
.venv/bin/alembic upgrade head
```

To inspect migration status:

```bash
.venv/bin/alembic current
.venv/bin/alembic history
```

## Start FastAPI

From `backend/`:

```bash
APP_ENV=development MPLCONFIGDIR=/tmp/matplotlib-report-app .venv/bin/python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

## Verify With API Calls

Create a report session:

```bash
curl -s -X POST http://127.0.0.1:8000/api/report-sessions \
  -H 'Content-Type: application/json' \
  -d '{
    "report_date": "2026-07-04",
    "station": "JUJA",
    "bound": "THIKA BOUND",
    "weighbridge_name": "JUJA",
    "prepared_by": "Tester",
    "confirmed_by": "Faith Njani"
  }'
```

Confirm metadata in PostgreSQL:

```bash
docker compose exec db psql -U report_app_user -d report_app_db \
  -c "select id, report_date, station, bound_name, status from reports order by updated_at desc limit 5;"
```

Inspect uploads/manual inputs after exercising the app:

```bash
docker compose exec db psql -U report_app_user -d report_app_db \
  -c "select report_id, upload_type, file_path from report_uploads order by uploaded_at desc limit 10;"

docker compose exec db psql -U report_app_user -d report_app_db \
  -c "select report_id, prepared_by, approved_by, traffic_total from report_manual_inputs order by updated_at desc limit 10;"
```

## Delete Bad or Mis-Uploaded Sessions

Report history and deletion are admin-only operations. The frontend exposes
them from the `/admin` page, which prompts for `ADMIN_PASSWORD` and sends it to
the backend as the `X-Admin-Password` header.

Admins can delete a session through the admin console, or directly through the
API:

```bash
curl -X DELETE http://127.0.0.1:8000/api/report-sessions/<report_id> \
  -H "X-Admin-Password: $ADMIN_PASSWORD"
```

This removes the session metadata, upload folders, processed DataFrames,
preview files, final outputs, in-memory session state, and PostgreSQL metadata
rows for that report.

## List Report History

Use the report-session listing endpoint to read saved report metadata from
PostgreSQL for a future report history or dashboard view:

```bash
curl 'http://127.0.0.1:8000/api/report-sessions?limit=20&offset=0' \
  -H "X-Admin-Password: $ADMIN_PASSWORD"
```

Optional query parameters:

- `status`: filter by report status, for example `draft` or `completed`
- `report_type`: filter by stored report type, for example `static_weighbridge`
- `limit`: number of sessions to return, default `20`, maximum `100`
- `offset`: pagination offset, default `0`
- `search`: simple search across title, weighbridge/station name, bound name,
  and report id

Results are sorted newest first by `created_at`.

Example response item:

```json
[
  {
    "report_id": "7f8f6b94-92df-4c2d-b1df-6a2428140474",
    "title": "JUJA - THIKA BOUND - 2026-07-04",
    "report_type": "static_weighbridge",
    "weighbridge_name": "JUJA",
    "bound_name": "THIKA BOUND",
    "status": "completed",
    "created_at": "2026-07-04T09:15:00Z",
    "updated_at": "2026-07-04T09:22:10Z",
    "completed_at": "2026-07-04T09:22:10Z",
    "has_final_report": true,
    "upload_count": 4,
    "required_uploads_completed": true,
    "manual_inputs_completed": true,
    "download_available": true
  }
]
```

The endpoint preserves the existing list response shape used by the frontend
and enriches each session with history fields. Internal file paths are not
exposed; `download_available` is calculated from report status, output metadata,
and the generated file's presence on disk when that check is available.

## Backfill Existing File Metadata Into PostgreSQL

If a deployment already has JSON report sessions on disk, use the backfill
script to copy existing metadata into PostgreSQL:

```bash
cd backend
.venv/bin/python scripts/backfill_report_metadata.py
```

The script persists:

- Report/session metadata into `reports`
- Manual input payloads into `report_manual_inputs`
- Upload file metadata into `report_uploads`
- Ready final output metadata into `report_outputs`

It does not copy file bytes into PostgreSQL. Uploaded files, processed
DataFrames, previews, and generated documents remain on the filesystem under
`REPORT_STORAGE_ROOT`.

## Manual Verification Checklist

- Create a report session and confirm a row appears in `reports`.
- Upload `daily_hour`, `wideload`, `impounded_prohibited`, and `overloaded`.
- Confirm upload paths appear in `report_uploads`.
- Save manual inputs and confirm a row appears in `report_manual_inputs`.
- Generate a section preview and confirm a row appears in `report_previews`.
- Build the final report and confirm `reports.status = completed`.
- Confirm generated file path appears in `report_outputs`.
- Download the final report and confirm the file still comes from filesystem storage.
- Delete a bad session with `X-Admin-Password` and confirm its files and
  database metadata are gone.
- Call `GET /api/report-sessions` with `X-Admin-Password` and confirm saved
  sessions are returned newest first with upload, manual-input, and download
  availability fields.
- Call `/api/report-sessions/analytics/dashboard` and confirm dashboard values
  reflect uploaded/static/mobile report data.

## Troubleshooting

### Production Persistence on Render

Check the deployed backend persistence status:

```bash
curl https://report-app-px6c.onrender.com/health/persistence
```

The expected production response has:

```json
{
  "status": "ok",
  "persistence_required": true,
  "database": {
    "configured": true,
    "connected": true,
    "error": null
  },
  "storage": {
    "configured": true,
    "root": "/var/data/report-app-storage"
  }
}
```

If `database.configured` is `false`, set `DATABASE_URL` on the Render web
service from the Render Postgres internal connection string. If
`storage.configured` is `false`, set `REPORT_STORAGE_ROOT` to:

```env
REPORT_STORAGE_ROOT=/var/data/report-app-storage
```

Also set:

```env
APP_ENV=production
ADMIN_PASSWORD=<your-admin-password>
```

The Render service must also have a persistent disk mounted at `/var/data`.
After changing environment variables, redeploy the service and run:

```bash
alembic upgrade head
```

The production health check uses `/health/persistence`, so Render will mark the
service unhealthy if PostgreSQL or persistent storage is not configured.

If `database.connected` is `false` with `database.error = "ProgrammingError"`,
the backend can reach PostgreSQL but the expected tables are missing. Run:

```bash
alembic upgrade head
```

The Docker startup command also runs `alembic upgrade head` before starting
Uvicorn so manually configured Render services still apply the schema on
deploy.

Current production verification completed on 2026-07-08:

```text
https://report-app-px6c.onrender.com/health/persistence
database.configured = true
database.connected = true
storage.root = /var/data/report-app-storage
```

A real JUJA static report was created through the production API and the
dashboard returned non-zero values, confirming new report metadata persists
through PostgreSQL and the analytics endpoint can read it.

Check database container:

```bash
docker compose ps
docker compose logs db
```

If you used `docker run`:

```bash
docker ps --filter name=report-app-postgres
docker logs report-app-postgres
```

Check connection:

```bash
docker compose exec db pg_isready -U report_app_user -d report_app_db
```

If you used `docker run`:

```bash
docker exec report-app-postgres pg_isready -U report_app_user -d report_app_db
```

Reset local database data:

```bash
docker compose down -v
docker compose up -d db
cd backend
.venv/bin/alembic upgrade head
```

If `alembic upgrade head` cannot find `DATABASE_URL`, confirm you are running
from `backend/` and that `.env` exists.
