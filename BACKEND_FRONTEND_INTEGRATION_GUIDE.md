# Backend and Frontend Integration Guide

This guide explains how the current backend works, how the report-generation files relate to each other, and how the frontend should be built so users can upload CSV/XLSX files, preview each report section as it is generated, and finally download one combined Word report.

The final goal is:

1. User enters report metadata such as report date, station, and bound.
2. User uploads the required CSV/XLSX files section by section.
3. Each upload is cleaned, processed, and previewed in the frontend.
4. The frontend progressively builds a preview of the final report as more sections become available.
5. User downloads the final `.docx` report with all sections combined and page count updated.

## Living Implementation Tracker

This document should be updated after every backend implementation step. Each time a route, service, preview flow, report section, or final-build behavior is completed, update this tracker before starting the next phase.

### Current Status Snapshot

```text
Status date: 2026-05-06
Backend framework: FastAPI
Frontend: Not started in this repo
Current backend mode: Section generators plus filesystem-backed report-session API
Current integration mode: Frontend-facing upload/preview/build/download routes are wired for implemented sections with restart recovery
```

### Completed

```text
FastAPI app entry point exists
Upload router exists
CSV/XLSX file reading exists for current upload endpoints
Shared template-based cleaning exists
Wideload standalone report generation exists
Impounded/prohibited standalone report generation exists
Standard report layout exists
Daily/hour raw data processing exists
Section 1 daily/hour statistics generator exists
Section 2 daily/hour table + graph generator exists
Section 2 has been refined to same-page table + chart layout
Final report builder exists in early form
Manual test script exists for daily/hour report generation
Manual test script exists for final report generation
Backend/frontend integration guide exists
Report-session store exists with filesystem-backed metadata and artifact persistence
Frontend-facing report-session routes exist
Daily/hour upload endpoint exists in report-session flow
Wideload upload endpoint exists in report-session flow
Impounded/prohibited upload endpoint exists in report-session flow
Overloaded upload endpoint exists in report-session flow
Report-session uploads are persisted under backend/app/storage/uploads/
Processed section DataFrames are persisted under backend/app/storage/processed/
Report-session metadata is persisted under backend/app/storage/sessions/
Generated previews are cached under backend/app/storage/previews/
Final reports are persisted under backend/app/storage/final_reports/
Report sessions can be reloaded after browser refresh or backend restart
Final report build endpoint exists for implemented sections
Final report download endpoint exists
Excel report download endpoint exists
Final report builder can include Sections 1, 2, 3, 4, 5, 6, and 7
Section preview renderer exists
Section preview endpoint exists
Preview URLs are returned for daily_hour, wideload, and impounded_prohibited uploads
Preview URL is returned for valid traffic_census manual input
Preview URL is returned for Daily Summary when required source data exists
Preview URL is returned for valid manual transgressions input
Preview endpoint supports png, pdf, and docx formats
Preview URL defaults to browser-friendly PNG output
Preview endpoint reuses cached DOCX/PDF/PNG artifacts when inputs have not changed
Manual report inputs can be stored on the report session
Prepared by and confirmed by can be rendered into the final report after Section 1
Traffic Census manual input is normalized and validated
Traffic Census preview rendering exists
Traffic Census final report rendering exists as Section 3
Daily Summary derivation exists from daily_hour, traffic_census, overloaded permit count, and manual inputs
Daily Summary preview rendering exists
Daily Summary final report rendering exists as Section 4
Transgressions manual input is normalized and validated
Transgressions preview rendering exists
Transgressions final report rendering exists as Section 5
Shared footer uses real Word PAGE and NUMPAGES fields
Automated API tests exist for persistent report-session recovery
Storage retention cleanup method exists for expired report-session artifacts
Reusable CSV fixtures exist for upload-through-build backend tests
Upload-through-build API test exists for report-session workflow
Upload-through-build API test covers Excel report download workbook structure
A4 landscape is the canonical final report page size
```

### In Progress

```text
Preparing backend hardening tasks after completing report rendering
```

### Pending

```text
Broaden fixture-based upload tests if more sample cases become available
```

### Implementation Log

Use this section as the running project history.

```text
2026-05-06 - Added backend/frontend integration guide.
2026-05-06 - Refined Section 2 daily-hour table + graph generator for same-page output.
2026-05-06 - Added living tracker and backend folder tree to this guide.
2026-05-06 - Added in-memory report-session store.
2026-05-06 - Added report-session API routes for create/get/upload/build/download.
2026-05-06 - Updated final report builder to include daily/hour Sections 1 and 2 before Sections 6 and 7.
2026-05-06 - Added section preview renderer and preview endpoint returning section-only .docx previews.
2026-05-06 - Upgraded preview endpoint to return PNG by default, with PDF and DOCX formats available by query parameter.
2026-05-06 - Added manual-input API support for prepared by, confirmed by, weighbridge name, traffic census, and transgressions.
2026-05-06 - Replaced the in-memory-only report-session store with filesystem-backed metadata, upload, processed-data, preview-cache, and final-report storage.
2026-05-06 - Added Traffic Census normalization, validation, preview rendering, and final report Section 3 rendering.
2026-05-06 - Added Daily Summary derivation, preview rendering, and final report Section 4 rendering.
2026-05-06 - Added Transgressions normalization, NIL rendering, preview rendering, and final report Section 5 rendering.
2026-05-06 - Replaced literal footer page placeholders with real Word PAGE and NUMPAGES fields.
2026-05-06 - Added pytest API tests for filesystem-backed report-session recovery, preview cache recovery, and final report persistence.
2026-05-06 - Added explicit storage cleanup for expired report-session metadata and artifacts.
2026-05-06 - Added reusable CSV fixtures and an upload-through-build API test for the report-session workflow.
2026-05-06 - Finalized A4 landscape as the canonical report page size and verified fixture-generated output.
2026-05-08 - Updated Traffic Census Section 3 so E derives from wideload upload count and Total Traffic derives from Q + X + K + E in previews and final DOCX builds.
2026-05-08 - Tuned final DOCX layout toward the Juja sample: Letter landscape page setup, sample table widths, larger Arial text, sample heading wording, and fewer forced page breaks.
2026-05-08 - Restored numbered section headings, corrected Traffic Census capitalization, added same-page section spacing, and restored wideload vertical-cell layout while keeping Letter page setup.
2026-05-08 - Added report-session summary-card endpoint for Total Weighed, Total Overloaded, Special Released, and Wide Loads; tuned E distribution window, wideload row height, and footer font size.
2026-05-08 - Finalized report-session backend for online handoff: upload/build flow, Section 1-7 previews, final DOCX output, wideload-derived E values, Letter layout, and frontend summary-card data are passing tests.
2026-05-11 - Added report-session Excel workbook generation and download endpoint using processed session data.
2026-05-11 - Updated Excel workbook generation to follow the Juja reference workbook layout, colors, dimensions, formulas, and embedded graph.
```

### Latest Backend Implementation Step

```text
Files changed:
- backend/app/services/excel_report_builder.py
- backend/app/routes/reports.py
- backend/tests/test_upload_build_flow.py
- BACKEND_FRONTEND_INTEGRATION_GUIDE.md

Behavior added:
- Report sessions now expose excel_report.download_url once daily_hour data is ready.
- GET /api/report-sessions/{report_id}/download-excel-report builds an XLSX workbook directly from processed session data.
- The workbook contains Summary and CC records sheets.
- Summary now follows the uploaded Juja workbook structure with fixed B:X table placement, matching row heights, column widths, merged cells, Arial fonts, yellow raw-data fills, red formula fills, and the daily/hour graph.
- CC records follows the reference summary layout with matching title/header placement, green total cells, and a NIL placeholder when detailed CC records are not available.

Rationale:
- The frontend needs an Excel download without converting the DOCX report.
- Generating the workbook from session data keeps the Excel feature separate from final DOCX generation while reusing the same processed values.

Tests and verification:
- Upload-through-build API test now downloads the Excel report, checks the Excel MIME type, opens it with openpyxl, and asserts required sheets, reference labels, row/column dimensions, fill colors, and chart presence.
- `backend/venv/bin/python -m compileall backend/app` passes.
- `PYTHONPATH=backend MPLCONFIGDIR=/tmp/matplotlib-cache backend/venv/bin/python -m pytest backend/tests -vv` passes.

Limitations:
- Detailed CC records are not captured by the current upload flow, so the CC records sheet emits a clear NIL placeholder for that table.
- Excel formatting now follows the reference workbook closely, but Excel and LibreOffice may render chart typography/spacing with small application-specific differences.

Next recommended task:
- Add a detailed CC records input/upload flow when the source data becomes available.
```

## Backend Folder Tree

Current backend layout:

```text
backend/
├── app/
│   ├── __init__.py
│   ├── main.py
│   ├── assets/
│   │   └── logo.png
│   ├── routes/
│   │   ├── __init__.py
│   │   ├── reports.py
│   │   └── upload.py
│   ├── services/
│   │   ├── __init__.py
│   │   ├── cleaner_core.py
│   │   ├── daily_hour_chart_generator.py
│   │   ├── daily_hour_generator.py
│   │   ├── daily_hour_processor.py
│   │   ├── daily_summary_generator.py
│   │   ├── daily_summary_processor.py
│   │   ├── excel_report_builder.py
│   │   ├── final_report_builder.py
│   │   ├── impounded_prohibited_generator.py
│   │   ├── overloaded_summary.py
│   │   ├── preview_renderer.py
│   │   ├── report_context.py
│   │   ├── report_layout.py
│   │   ├── report_session_store.py
│   │   ├── report_upload_service.py
│   │   ├── traffic_census_generator.py
│   │   ├── traffic_census_processor.py
│   │   ├── transgressions_generator.py
│   │   ├── transgressions_processor.py
│   │   └── wideload_generator.py
│   ├── storage/
│   │   ├── final_reports/
│   │   ├── previews/
│   │   ├── processed/
│   │   ├── sessions/
│   │   └── uploads/
│   └── templates/
│       ├── __init__.py
│       ├── impounded_prohibited.py
│       └── vehicle_inspection.py
├── test_daily_hour_report.py
├── test_final_report.py
├── tests/
│   ├── conftest.py
│   ├── fixtures/
│   │   ├── daily_hour.csv
│   │   ├── impounded_prohibited.csv
│   │   ├── overloaded.csv
│   │   └── wideload.csv
│   ├── test_report_session_api.py
│   ├── test_storage_cleanup.py
│   └── test_upload_build_flow.py
├── daily_hour_test.docx
├── final_report_test.docx
└── venv/
```

Ignored/generated folders and files:

```text
__pycache__/
*.pyc
*.docx
.~lock.*
backend/venv/
backend/app/storage/
```

Recommended future backend layout:

```text
backend/
├── app/
│   ├── main.py
│   ├── assets/
│   ├── routes/
│   │   ├── upload.py
│   │   └── reports.py
│   ├── schemas/
│   │   ├── report_session.py
│   │   └── section_status.py
│   ├── services/
│   │   ├── cleaner_core.py
│   │   ├── report_session_store.py
│   │   ├── preview_renderer.py
│   │   ├── final_report_builder.py
│   │   ├── daily_hour_processor.py
│   │   ├── daily_hour_generator.py
│   │   ├── daily_hour_chart_generator.py
│   │   ├── wideload_generator.py
│   │   ├── impounded_prohibited_generator.py
│   │   ├── overloaded_summary.py
│   │   ├── traffic_census_processor.py
│   │   ├── traffic_census_generator.py
│   │   ├── daily_summary_generator.py
│   │   ├── transgressions_processor.py
│   │   └── transgressions_generator.py
│   ├── templates/
│   └── storage/
│       ├── uploads/
│       ├── previews/
│       └── final_reports/
└── tests/
    ├── test_daily_hour_report.py
    ├── test_final_report.py
    └── test_report_session_api.py
```

The future layout is a guide, not a requirement to create every file immediately. We should add folders only when an implementation step needs them.

## Current Backend Shape

The backend is a FastAPI application.

Entry point:

```text
backend/app/main.py
```

The app includes routes from:

```text
backend/app/routes/upload.py
```

Current API routes:

```text
POST /api/process-file
POST /api/download-wideload-report
POST /api/download-impounded-prohibited-report
POST /api/report-sessions
GET /api/report-sessions/{report_id}
PATCH /api/report-sessions/{report_id}/manual-inputs
POST /api/report-sessions/{report_id}/uploads/daily-hour
POST /api/report-sessions/{report_id}/uploads/wideload
POST /api/report-sessions/{report_id}/uploads/impounded-prohibited
POST /api/report-sessions/{report_id}/uploads/overloaded
POST /api/report-sessions/{report_id}/build-final-report
GET /api/report-sessions/{report_id}/sections/{section_name}/preview
GET /api/report-sessions/{report_id}/download-final-report
GET /api/report-sessions/{report_id}/download-excel-report
```

At the moment, the backend can clean uploaded data, generate standalone Word reports for wideload and impounded/prohibited sections, and create filesystem-backed report sessions for the frontend workflow. Daily/hour data, wideload data, impounded/prohibited data, and overloaded data can now be uploaded into a report session and recovered from disk by `report_id`. Manual report inputs can be stored for prepared by, confirmed by, weighbridge name, traffic census, cases cleared in court, transgressions count, and transgressions. Traffic Census manual input is validated and can be rendered as Section 3. Daily Summary is derived from ready source sections and can be rendered as Section 4. Transgressions manual input is normalized and can be rendered as Section 5 with NIL rows when either table has no records. Section previews are available for daily/hour, traffic census, daily summary, transgressions, wideload, and impounded/prohibited sections as cached PNG, PDF, or DOCX artifacts. The final report can be built, persisted, recovered, and downloaded for the currently implemented sections. The Excel report can be downloaded directly from processed session data when daily/hour data is ready.

## Current File Responsibilities

### `backend/app/main.py`

Creates the FastAPI app and registers the upload router.

Purpose:

```text
FastAPI app startup
Router registration
```

### `backend/app/routes/upload.py`

Currently handles frontend/API upload requests.

Implemented:

```text
POST /api/process-file
POST /api/download-wideload-report
POST /api/download-impounded-prohibited-report
```

This file now remains as the legacy/standalone upload route module. New frontend report-builder work should use `reports.py`.

### `backend/app/routes/reports.py`

Handles report-session API requests for the frontend workflow.

Implemented:

```text
POST /api/report-sessions
GET /api/report-sessions/{report_id}
PATCH /api/report-sessions/{report_id}/manual-inputs
POST /api/report-sessions/{report_id}/uploads/daily-hour
POST /api/report-sessions/{report_id}/uploads/wideload
POST /api/report-sessions/{report_id}/uploads/impounded-prohibited
POST /api/report-sessions/{report_id}/uploads/overloaded
POST /api/report-sessions/{report_id}/build-final-report
GET /api/report-sessions/{report_id}/sections/{section_name}/preview
GET /api/report-sessions/{report_id}/download-final-report
GET /api/report-sessions/{report_id}/download-excel-report
```

This module stores report state through `report_session_store.py`. The API response shape remains session-oriented, while metadata and generated artifacts are now persisted under `backend/app/storage/` so sessions can be recovered by `report_id`.

### `backend/app/services/preview_renderer.py`

Builds section-only preview documents from report-session data.

Currently supports:

```text
daily_hour
traffic_census
daily_summary
transgressions
wideload
impounded_prohibited
```

Supported preview formats:

```text
png
pdf
docx
```

The preview route defaults to PNG for browser display:

```text
GET /api/report-sessions/{report_id}/sections/{section_name}/preview
```

Explicit format examples:

```text
GET /api/report-sessions/{report_id}/sections/{section_name}/preview?format=png
GET /api/report-sessions/{report_id}/sections/{section_name}/preview?format=pdf
GET /api/report-sessions/{report_id}/sections/{section_name}/preview?format=docx
```

Daily/hour previews have two pages:

```text
page=1 -> Daily and Hourly Statistics
page=2 -> Daily Hourly Data table + chart
```

Generated preview artifacts are cached under:

```text
backend/app/storage/previews/{report_id}/
```

When a section upload or manual input changes, cached previews for that session are invalidated.

### `backend/app/services/cleaner_core.py`

Shared cleaning engine.

It:

```text
Validates required columns
Selects required columns
Renames raw CSV/XLSX columns into report columns
Formats date columns
Formats comma-number columns
Returns cleaned pandas DataFrame
```

It depends on template modules that define column mappings.

### `backend/app/templates/vehicle_inspection.py`

Defines the raw input columns and output order for wideload/vehicle inspection data.

Used by:

```text
clean_with_template(df, vehicle_inspection)
```

### `backend/app/templates/impounded_prohibited.py`

Defines the raw input columns and output order for impounded/prohibited data.

Used by:

```text
clean_with_template(df, impounded_prohibited)
```

### `backend/app/services/report_layout.py`

Applies the shared Word document layout.

It currently handles:

```text
Canonical A4 landscape page setup
Margins
Header logo
Footer text
Real Word PAGE and NUMPAGES footer fields
Report date formatting
Shared A4 printable width and table width constants
```

All final report sections should be added to a document after this standard layout has been applied.

Page numbering is implemented with real Word `PAGE` and `NUMPAGES` fields. The document settings request field updates when the DOCX is opened.

Canonical page-size policy:

```text
Final report page size: A4 landscape
Page width: 11.69 inches
Page height: 8.27 inches
Left/right margins: 0.25 inches
Top margin: 0.75 inches
Bottom margin: 0.55 inches
Shared wide-table width: 16,000 twips
```

Rationale:

```text
The backend already used A4 landscape as its shared layout.
Existing sections fit A4 landscape with the configured tight margins.
Using one canonical A4 layout avoids mixing Letter-specific assumptions into table and chart generation.
```

### `backend/app/services/report_session_store.py`

Filesystem-backed report-session storage for the frontend-facing backend.

It stores:

```text
Report metadata
Section statuses
Uploaded/processed DataFrames
Final report bytes
Final report status/error
```

Persistent paths:

```text
backend/app/storage/sessions/
backend/app/storage/uploads/
backend/app/storage/processed/
backend/app/storage/previews/
backend/app/storage/final_reports/
```

The store keeps the existing in-process session behavior, but it now writes each mutation to disk and lazily reloads sessions from disk when needed after backend restart.

Storage cleanup:

```text
Default retention window: 168 hours / 7 days
cleanup_expired_sessions(max_age_hours=168) removes expired sessions and related artifacts
Cleanup is explicit maintenance behavior and is not run automatically on each API request
```

### `backend/app/services/report_upload_service.py`

Reads FastAPI upload files into bytes and pandas DataFrames.

Purpose:

```text
Central CSV/XLSX upload parsing for report-session routes
Return raw bytes so uploads can be persisted before processing
Keep file-reading details out of route handlers
```

### `backend/app/services/traffic_census_processor.py`

Normalizes and validates manual Traffic Census input.

Expected input:

```text
buses_gte_3500kg
vehicles_3500_to_7000_excluding_buses
vehicles_gte_7000_excluding_buses
total_traffic_census
```

If `total_traffic_census` is omitted, the processor computes it from the three category fields. If it is provided and does not match the computed total, the processor raises a validation error.

### `backend/app/services/traffic_census_generator.py`

Generates Section 3:

```text
3. TRAFFIC CENSUS DATA
```

Main function:

```text
add_traffic_census_section(doc, traffic_census_data)
```

This section uses normalized manual input from the report session.

### `backend/app/services/wideload_generator.py`

Generates the wideload/vehicle inspection section.

Provides:

```text
add_wideload_section(doc, df)
generate_wideload_report(df, report_date, station, bound)
```

`add_wideload_section` is the section-level function that should be used by the final combined report.

`generate_wideload_report` creates a standalone report document.

### `backend/app/services/impounded_prohibited_generator.py`

Generates the impounded/prohibited section.

Provides:

```text
add_impounded_prohibited_section(doc, df)
generate_impounded_prohibited_report(df, report_date, station, bound)
```

`add_impounded_prohibited_section` should be used by the final combined report.

`generate_impounded_prohibited_report` creates a standalone report document.

### `backend/app/services/daily_hour_processor.py`

Processes raw daily/hour statistics data.

It:

```text
Validates required daily/hour columns
Extracts core numeric columns
Computes D, S, M, H, Q, X, C, Y, P, A, Z, G, R, E
Distributes wideload exemption permits across daytime hours
Adds the totals row
```

Main functions:

```text
build_daily_hour_metrics(raw_df, report_date, wideload_count)
add_daily_totals_row(daily_df)
```

### `backend/app/services/daily_hour_generator.py`

Generates Section 1:

```text
1. DAILY AND HOURLY STATISTICS
```

Main function:

```text
add_daily_hour_statistics_section(doc, daily_df)
```

This section uses the processed `daily_df` from `daily_hour_processor.py`.

### `backend/app/services/daily_hour_chart_generator.py`

Generates Section 2:

```text
2. DAILY HOURLY DATA
```

It builds:

```text
Compact hourly data table
Line chart image
Same-page table + chart layout
```

Main functions:

```text
build_daily_hour_chart_data(daily_df)
create_daily_hour_chart_image(chart_df)
add_daily_hour_chart_section(doc, daily_df)
```

This section depends on the already-processed `daily_df`.

### `backend/app/services/daily_summary_processor.py`

Derives Section 4 values from existing report-session data.

Required sources:

```text
daily_hour processed totals
traffic_census manual input
overloaded valid_permit_count
```

Manual/default fields:

```text
cases_cleared_in_court -> default 0
transgressions_count / total_transgressions -> default 0
```

Daily Summary formula note:

```text
Impounded & prohibited (P) = Z + R + G
```

This is intentionally different from the Section 1 daily/hour table formula, where `P = Z + R`.

Main functions:

```text
build_daily_summary(...)
build_daily_summary_from_session(session)
daily_summary_rows(summary)
```

### `backend/app/services/daily_summary_generator.py`

Generates Section 4:

```text
4. DAILY SUMMARY
```

Main function:

```text
add_daily_summary_section(doc, summary)
```

This section uses the derived summary dict from `daily_summary_processor.py`.

### `backend/app/services/transgressions_processor.py`

Normalizes manual Transgressions input into two report-ready lists.

Expected input:

```text
transgressions.daily_transgressions
transgressions.action_report
```

Supported behavior:

```text
Empty lists are valid and render as NIL rows
Common key aliases are mapped into report column names
```

### `backend/app/services/transgressions_generator.py`

Generates Section 5:

```text
5. TRANSGRESSIONS
```

It renders:

```text
DAILY TRANSGRESSIONS REPORT
I. TRANSGRESSIONS ACTION REPORT
```

Main function:

```text
add_transgressions_section(doc, transgressions_data)
```

### `backend/app/services/overloaded_summary.py`

Counts valid permit vehicles from overloaded data.

Main function:

```text
count_valid_permit_vehicles(overloaded_df)
```

Currently used by the final report builder to calculate permit-related summary values.

### `backend/app/services/final_report_builder.py`

Builds a combined report document.

Current function:

```text
build_final_report(
    wideload_df,
    impounded_prohibited_df,
    overloaded_df,
    report_date,
    station,
    bound,
    traffic_census=None,
    daily_summary=None,
    transgressions=None,
)
```

Current implementation combines:

```text
Daily/hour statistics when provided
Daily/hour chart when provided
Traffic Census when valid manual input is provided
Daily Summary when derived source values are ready
Transgressions when manual input is provided
Impounded/prohibited section
Wideload section
```

It does not yet include:

```text
Daily summary
Transgressions
All final report sections in correct order
```

This file should become the main backend assembly point for the final downloadable report.

## How To Test The Backend Now

Run the backend from the `backend/` directory:

```bash
cd backend
venv/bin/uvicorn app.main:app --reload
```

The API will be available at:

```text
http://127.0.0.1:8000
```

Interactive docs:

```text
http://127.0.0.1:8000/docs
```

Run automated backend API tests from the repo root:

```bash
PYTHONPATH=backend MPLCONFIGDIR=/tmp/matplotlib-cache backend/venv/bin/python -m pytest backend/tests
```

The API tests use temporary filesystem storage, so they do not write to `backend/app/storage/`.

### 1. Create A Report Session

Frontend endpoint:

```text
POST /api/report-sessions
```

Example request body:

```json
{
  "report_date": "2026-02-02",
  "station": "Juja",
  "bound": "Thika Bound",
  "weighbridge_name": "Juja",
  "prepared_by": "Fredrick Kariuki",
  "confirmed_by": "Faith Njani"
}
```

`station` and `weighbridge_name` are interchangeable at creation time. The frontend can send only `weighbridge_name` if that is the label used in the UI.

Example response shape:

```json
{
  "report_id": "uuid",
  "metadata": {
    "report_date": "2026-02-02",
    "station": "Juja",
    "bound": "Thika Bound",
    "weighbridge_name": "Juja",
    "prepared_by": "Fredrick Kariuki",
    "confirmed_by": "Faith Njani"
  },
  "sections": {
    "daily_hour": { "status": "missing" },
    "wideload": { "status": "missing" },
    "impounded_prohibited": { "status": "missing" },
    "overloaded": { "status": "missing" }
  },
  "final_report": {
    "status": "not_built",
    "download_url": null,
    "error": null
  },
  "excel_report": {
    "status": "awaiting_data",
    "download_url": null
  }
}
```

The frontend should store `report_id`.

### 2. Update Manual Inputs

Frontend endpoint:

```text
PATCH /api/report-sessions/{report_id}/manual-inputs
```

Example request body:

```json
{
  "prepared_by": "Fredrick Kariuki",
  "confirmed_by": "Faith Njani",
  "weighbridge_name": "Juja",
  "traffic_census": {
    "buses_gte_3500kg": 1351,
    "vehicles_3500_to_7000_excluding_buses": 29,
    "vehicles_gte_7000_excluding_buses": 6,
    "total_traffic_census": 1386
  },
  "transgressions": {
    "daily_transgressions": [],
    "action_report": []
  }
}
```

Current behavior:

```text
prepared_by and confirmed_by are rendered into the final report after Section 1.
traffic_census is normalized, validated, stored, and rendered as Section 3 when valid.
transgressions is normalized, stored, and rendered as Section 5 when provided.
```

### 3. Upload Daily Hour Statistics

Frontend endpoint:

```text
POST /api/report-sessions/{report_id}/uploads/daily-hour
```

Form data:

```text
file: daily hour CSV/XLSX
wideload_count: number
```

Example curl:

```bash
curl -X POST \
  -F "file=@/path/to/Daily Hour Statistics.csv" \
  -F "wideload_count=25" \
  http://127.0.0.1:8000/api/report-sessions/{report_id}/uploads/daily-hour
```

Response includes:

```json
{
  "sections": {
    "daily_hour": {
      "status": "ready",
      "preview_url": "/api/report-sessions/{report_id}/sections/daily_hour/preview?format=png",
      "preview_pages": [
        {
          "label": "Daily and Hourly Statistics",
          "url": "/api/report-sessions/{report_id}/sections/daily-hour-statistics/preview?format=png&page=1"
        },
        {
          "label": "Daily Hourly Data",
          "url": "/api/report-sessions/{report_id}/sections/daily-hour-chart/preview?format=png&page=2"
        }
      ]
    }
  }
}
```

### 4. Upload Wideload File

Frontend endpoint:

```text
POST /api/report-sessions/{report_id}/uploads/wideload
```

Form data:

```text
file: wideload CSV/XLSX
```

Preview URL returned:

```text
/api/report-sessions/{report_id}/sections/wideload/preview?format=png
```

### 5. Upload Impounded/Prohibited File

Frontend endpoint:

```text
POST /api/report-sessions/{report_id}/uploads/impounded-prohibited
```

Form data:

```text
file: impounded/prohibited CSV/XLSX
```

Preview URL returned:

```text
/api/report-sessions/{report_id}/sections/impounded_prohibited/preview?format=png
```

### 6. Upload Overloaded File

Frontend endpoint:

```text
POST /api/report-sessions/{report_id}/uploads/overloaded
```

Form data:

```text
file: overloaded CSV/XLSX
```

Current behavior:

```text
The backend stores the raw overloaded DataFrame.
The backend computes valid_permit_count using count_valid_permit_vehicles().
No visual section preview is generated for overloaded data yet.
```

### 7. Preview Sections

Frontend can render PNG previews directly in an image tag.

Example:

```html
<img src="http://127.0.0.1:8000/api/report-sessions/{report_id}/sections/daily-hour-chart/preview?format=png&page=2" />
```

Supported formats:

```text
format=png
format=pdf
format=docx
```

Examples:

```text
GET /api/report-sessions/{report_id}/sections/daily-hour-statistics/preview?format=png&page=1
GET /api/report-sessions/{report_id}/sections/daily-hour-chart/preview?format=png&page=2
GET /api/report-sessions/{report_id}/sections/traffic-census/preview?format=png
GET /api/report-sessions/{report_id}/sections/daily-summary/preview?format=png
GET /api/report-sessions/{report_id}/sections/transgressions/preview?format=png
GET /api/report-sessions/{report_id}/sections/wideload/preview?format=pdf
GET /api/report-sessions/{report_id}/sections/impounded-prohibited/preview?format=docx
```

### 8. Build Final Report

Frontend endpoint:

```text
POST /api/report-sessions/{report_id}/build-final-report
```

Required ready sections:

```text
daily_hour
wideload
impounded_prohibited
overloaded
```

If anything is missing, response shape:

```json
{
  "final_report": {
    "status": "error",
    "download_url": null,
    "error": "Missing or invalid required sections: [...]"
  }
}
```

If build succeeds:

```json
{
  "final_report": {
    "status": "ready",
    "download_url": "/api/report-sessions/{report_id}/download-final-report",
    "error": null
  }
}
```

### 9. Download Final Report

Frontend endpoint:

```text
GET /api/report-sessions/{report_id}/download-final-report
```

The frontend can use this URL as the download link after `final_report.status === "ready"`.

Current final report includes:

```text
Section 1: Daily and Hourly Statistics
Prepared by / Confirmed by lines
Section 2: Daily Hourly Data table + graph
Section 3: Traffic Census Data, when valid manual traffic_census input exists
Section 4: Daily Summary, when required source data exists
Section 5: Transgressions, when manual transgressions input exists
Section 6: Impounded & Prohibited
Section 7: Vehicle Inspection Report (Wide Loads)
```

Current final report visual rendering status:

```text
All target sections are now visually renderable when their source data exists
A4 landscape is the canonical final report page size
Fixture-generated report output has been visually inspected after DOCX-to-PDF conversion
```

### 10. Download Excel Report

Frontend endpoint:

```text
GET /api/report-sessions/{report_id}/download-excel-report
```

The frontend can use:

```js
`${API_BASE}/api/report-sessions/${reportId}/download-excel-report`
```

The serialized session also includes:

```json
{
  "excel_report": {
    "status": "ready",
    "download_url": "/api/report-sessions/{report_id}/download-excel-report"
  }
}
```

Current Excel workbook includes:

```text
Sheet 1: Summary
- Daily and Hourly Statistics
- Daily Hour Data
- Traffic Census Data
- Daily Summary

Sheet 2: CC records
- Traffic census summary
- Census / CC records table with a NIL placeholder until detailed CC records are captured
```

Remaining report work is polish and hardening, including broader fixture coverage as more sample cases become available.

## Desired Final Report Flow

The intended final report should be assembled in this order:

```text
1. Daily and Hourly Statistics
2. Daily Hourly Data
3. Traffic Census Data
4. Daily Summary
5. Transgressions
6. Impounded & Prohibited
7. Vehicle Inspection Report (Wide Loads)
```

Transgressions is manual-input driven for now. A CSV/XLSX upload flow can be added later if needed.

## Recommended Backend Architecture for Frontend Integration

The frontend should not call many standalone download endpoints and try to combine documents itself. The backend should own report generation because:

```text
Word layout is backend-specific
Page breaks must be controlled centrally
Footer/page count must be generated consistently
Derived values depend on multiple files
The final report should be assembled in one place
```

Recommended backend concept:

```text
Report Session
```

A report session stores:

```text
report_id
report_date
station
bound
uploaded raw files
cleaned DataFrames or serialized cleaned data
section generation status
section preview files
final report status
final report file
errors per section
```

For local/simple development, session state can start in memory or in `/tmp`. For production, use persistent storage:

```text
Database for metadata
Object/file storage for uploaded CSV/XLSX and generated docx/pdf/png previews
Background worker for long report generation jobs
```

## Recommended API Endpoints

### Create Report Session

```http
POST /api/report-sessions
```

Request:

```json
{
  "report_date": "2026-02-02",
  "station": "Juja",
  "bound": "Thika Bound"
}
```

Response:

```json
{
  "report_id": "uuid",
  "status": "created"
}
```

### Upload Daily Hour Statistics File

```http
POST /api/report-sessions/{report_id}/uploads/daily-hour
```

Form data:

```text
file: CSV/XLSX
wideload_count: number
```

Backend should:

```text
Read CSV/XLSX
Run build_daily_hour_metrics(...)
Run add_daily_totals_row(...)
Generate Section 1 preview
Generate Section 2 preview
Store cleaned daily_df for final report assembly
Return section statuses
```

Response:

```json
{
  "report_id": "uuid",
  "sections": {
    "daily_hour_statistics": {
      "status": "ready",
      "preview_url": "/api/report-sessions/uuid/sections/daily-hour-statistics/preview"
    },
    "daily_hour_chart": {
      "status": "ready",
      "preview_url": "/api/report-sessions/uuid/sections/daily-hour-chart/preview"
    }
  }
}
```

### Upload Wideload File

```http
POST /api/report-sessions/{report_id}/uploads/wideload
```

Form data:

```text
file: CSV/XLSX
```

Backend should:

```text
Clean file with vehicle_inspection template
Store cleaned wideload_df
Generate wideload section preview
Update wideload_count used by daily/hour E distribution if needed
```

Important: daily/hour processing currently accepts `wideload_count`. The final system should decide whether this count comes from:

```text
Uploaded wideload file row count
Manual frontend field
Both, with manual override
```

Recommended: derive it from the cleaned wideload file by default, then allow a manual override only if necessary.

### Upload Impounded/Prohibited File

```http
POST /api/report-sessions/{report_id}/uploads/impounded-prohibited
```

Backend should:

```text
Clean file with impounded_prohibited template
Store cleaned impounded_prohibited_df
Generate impounded/prohibited section preview
```

### Upload Overloaded File

```http
POST /api/report-sessions/{report_id}/uploads/overloaded
```

Backend should:

```text
Read raw overloaded data
Run count_valid_permit_vehicles(...)
Store overloaded_df
Store permit summary values for final report
```

### Get Report Session Status

```http
GET /api/report-sessions/{report_id}
```

Response:

```json
{
  "report_id": "uuid",
  "metadata": {
    "report_date": "2026-02-02",
    "station": "Juja",
    "bound": "Thika Bound"
  },
  "sections": {
    "daily_hour_statistics": "ready",
    "daily_hour_chart": "ready",
    "traffic_census": "missing",
    "daily_summary": "missing",
    "transgressions": "missing",
    "impounded_prohibited": "ready",
    "wideload": "ready"
  },
  "final_report": {
    "status": "not_built",
    "download_url": null
  }
}
```

### Get Section Preview

```http
GET /api/report-sessions/{report_id}/sections/{section_name}/preview
```

Recommended preview formats:

```text
PDF preview
PNG page image preview
HTML table/chart preview
```

Best option for visual fidelity:

```text
Generate section docx
Convert section docx to PDF
Convert PDF page to PNG
Return PNG/PDF preview URL to frontend
```

This gives the frontend a close visual match to the Word document.

### Build Final Report

```http
POST /api/report-sessions/{report_id}/build-final-report
```

Backend should:

```text
Validate all required uploads exist
Build one Document()
Apply standard layout once
Add all implemented sections in order
Insert page breaks where required
Save final docx
Return status and download URL
```

Response:

```json
{
  "report_id": "uuid",
  "status": "building"
}
```

If generation is fast, this can return ready immediately:

```json
{
  "report_id": "uuid",
  "status": "ready",
  "download_url": "/api/report-sessions/uuid/download-final-report"
}
```

### Download Final Report

```http
GET /api/report-sessions/{report_id}/download-final-report
```

Response:

```text
application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

Filename example:

```text
juja_thika_bound_daily_report_2026-02-02.docx
```

## Preview Generation Strategy

The frontend wants to show previews as each file is uploaded.

There are two practical preview approaches.

### Option 1: Backend-rendered preview images

Backend generates each section as a temporary `.docx`, converts it to PDF, then converts the PDF page to PNG.

Pros:

```text
Most visually accurate
Frontend stays simple
Matches final Word layout better
Easy to display in browser
```

Cons:

```text
Requires LibreOffice on backend
Requires PDF/image conversion tool such as poppler
Preview generation may need background jobs
```

Recommended for this project because the goal is precise Word report formatting.

### Option 2: Frontend HTML preview

Backend returns cleaned JSON/table data and chart data. Frontend renders preview using HTML tables and a chart library.

Pros:

```text
Fast
Interactive
No server-side document conversion needed
```

Cons:

```text
Will not perfectly match final Word layout
Frontend must duplicate layout logic
Chart/table may look different from final docx
```

Recommended only for quick data preview, not final visual preview.

### Best Combined Approach

Use both:

```text
Fast JSON preview immediately after upload
Backend-rendered image/PDF preview when ready
```

Frontend behavior:

```text
Show "Processing..." after upload
Show cleaned data preview quickly
Replace or supplement with rendered section preview when backend finishes
```

## Frontend Build Guide

The frontend should be built around a report builder workflow, not independent upload pages.

Recommended main screens/components:

```text
ReportMetadataForm
UploadChecklist
SectionUploadCard
SectionPreviewPanel
FinalReportPreviewPanel
BuildReportButton
DownloadReportButton
ErrorPanel
```

### Frontend State Model

Recommended state shape:

```ts
type SectionStatus = "missing" | "uploading" | "processing" | "ready" | "error" | "not_implemented";

type ReportBuilderState = {
  reportId: string | null;
  metadata: {
    reportDate: string;
    station: string;
    bound: string;
  };
  sections: {
    dailyHourStatistics: SectionState;
    dailyHourChart: SectionState;
    trafficCensus: SectionState;
    dailySummary: SectionState;
    transgressions: SectionState;
    impoundedProhibited: SectionState;
    wideload: SectionState;
    overloaded: SectionState;
  };
  finalReport: {
    status: "not_ready" | "ready_to_build" | "building" | "ready" | "error";
    previewUrl?: string;
    downloadUrl?: string;
    error?: string;
  };
};

type SectionState = {
  status: SectionStatus;
  uploadName?: string;
  previewUrl?: string;
  jsonPreview?: unknown;
  error?: string;
};
```

### Frontend Upload Flow

1. User fills report metadata.
2. Frontend calls `POST /api/report-sessions`.
3. Backend returns `report_id`.
4. User uploads each required file.
5. Frontend sends file to the matching upload endpoint.
6. Frontend marks section as `uploading`.
7. Backend starts processing.
8. Frontend marks section as `processing`.
9. Frontend polls `GET /api/report-sessions/{report_id}` or listens via websocket/server-sent events.
10. When section status becomes `ready`, frontend displays preview URL.
11. Once all required sections are ready, frontend enables `Build Final Report`.
12. User clicks build.
13. Frontend calls `POST /api/report-sessions/{report_id}/build-final-report`.
14. Frontend polls until final report status is `ready`.
15. Frontend enables `Download Final Report`.

### Required Uploads for Current Build

For the currently implemented backend, the frontend should support:

```text
Daily Hour Statistics CSV/XLSX
Wideload / Vehicle Inspection CSV/XLSX
Impounded & Prohibited CSV/XLSX
Overloaded CSV/XLSX
```

Future uploads:

```text
Traffic Census CSV/XLSX
Optional Transgressions CSV/XLSX if manual entry is not enough
Any manual summary fields needed by Daily Summary
```

## How Final Report Assembly Should Work

The final builder should not use standalone generator functions like `generate_wideload_report` because those create separate documents.

Instead it should use section functions:

```python
doc = Document()
apply_standard_layout(doc, report_date, station, bound)

add_daily_hour_statistics_section(doc, daily_df)
doc.add_page_break()

add_daily_hour_chart_section(doc, daily_df)
doc.add_page_break()

add_traffic_census_section(doc, traffic_census_data)
doc.add_page_break()

add_daily_summary_section(doc, summary)
doc.add_page_break()

add_transgressions_section(doc, transgressions_df)
doc.add_page_break()

add_impounded_prohibited_section(doc, impounded_prohibited_df)
doc.add_page_break()

add_wideload_section(doc, wideload_df)
```

The exact page breaks should follow the target report design.

## Backend Work Still Needed

To fully harden the frontend workflow, implement the following:

```text
Broaden fixture-based upload tests if more sample cases become available
```

## Suggested Development Phases

### Phase 1: Stabilize Backend Report Sections

Goal:

```text
Each implemented section can be generated from a function that accepts clean data and appends to an existing Document.
```

Tasks:

```text
Keep add_*_section functions as the main reusable API
Avoid duplicating layout code inside section generators
Confirm Section 1 and Section 2 visual output
Confirm Section 6 and Section 7 visual output against target document
```

### Phase 2: Add Backend Report Session API

Goal:

```text
Frontend can create one report session and upload files into it.
```

Tasks:

```text
Create report session endpoints
Store uploaded files and cleaned data
Return section status JSON
Return useful validation errors
```

### Phase 3: Add Preview Generation

Goal:

```text
Frontend can show preview after each upload.
```

Tasks:

```text
Generate section-only docx previews
Convert previews to PDF/PNG
Expose preview URLs
Add frontend polling
```

### Phase 4: Build Final Report Endpoint

Goal:

```text
Frontend can request the final combined report and download it.
```

Tasks:

```text
Update final_report_builder.py to include all implemented sections
Validate missing required files before build
Return downloadable final .docx
Update page fields
```

### Phase 5: Add Missing Sections

Goal:

```text
Complete full target report.
```

Tasks:

```text
Any manual input forms required by sections that are not fully CSV-driven
```

## Error Handling Requirements

The backend should return clear errors for:

```text
Unsupported file format
Missing required columns
Invalid dates
Empty files
Failed section generation
Failed preview generation
Missing required uploads before final build
```

Example response:

```json
{
  "status": "error",
  "section": "daily_hour_statistics",
  "message": "Missing columns: ['MultiDeck[D]', 'SingleAxle[S]']"
}
```

The frontend should show these errors inside the matching upload card.

## Open Decisions

These should be clarified before building the full frontend:

1. Should `wideload_count` always come from the uploaded wideload file, or should users be allowed to type/override it?
2. Should the frontend display cached PNG previews, embedded PDFs, or both? The backend supports cached PNG, PDF, and DOCX.
3. Is the current 7-day local filesystem retention window appropriate for production?
4. Should generated reports be stored permanently, or deleted after a time window?
5. Which sections are mandatory before the final download button is enabled?
6. Should Transgressions remain manual-input driven, or should a CSV/XLSX upload flow be added later?

Resolved decisions:

```text
Final report page size is A4 landscape.
Letter landscape is no longer treated as an active backend output option.
```

## Recommended Immediate Next Step

Broaden fixture-based upload tests if more sample cases become available.

Target files:

```text
backend/tests/fixtures/
backend/tests/test_upload_build_flow.py
BACKEND_FRONTEND_INTEGRATION_GUIDE.md
```

Current behavior:

```text
Small reusable CSV fixtures cover the happy-path upload-through-build flow
The upload-through-build API test verifies Letter landscape page dimensions and table grid widths
Section 1-7 previews and final DOCX generation are available through report-session routes
Summary cards endpoint is available for Total Weighed, Total Overloaded, Special Released, and Wide Loads
```

Recommended next behavior:

```text
Add more realistic fixture cases with longer vehicle names, routes, permit values, and transgression rows
Use those fixtures to tune row heights, font sizes, and column ratios if needed
Host the backend and connect the frontend to the report-session API routes
```

After broader fixtures are stable, continue with:

```text
Frontend report-builder implementation planning
```
