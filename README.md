# NAOS SLA Tracker

Desktop (PySide6) + CLI tool to ingest Outlook mail, track SLA in SQLite, and sync/export to Excel. Classic Outlook COM only; SAFE mode by default, SEND must be explicit. Developer: Tyoma (ShapArt).

## Quick start

1. Python 3.11+, Windows with Classic Outlook installed.
2. `python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt`
3. Run CLI: `python cli.py --help`
   - `python cli.py ingest --days 7`
   - `python cli.py recalc-open`
   - `python cli.py export-xlsx`
   - `python cli.py process-responses`
   - `python cli.py send-overdue` (preview if SAFE)
   - `python cli.py sync-all` (full pipeline Excel->DB->Outlook ingest->recalc->export->notify preview)
4. Run UI: `python cli.py ui` or launch `launch_ui.py`. Toggle SAFE/SEND in the settings panel; SEND asks for confirmation and still obeys allowlist/domains.

## Safety & sending

- `safe_mode=True` and `allow_send=False` by default.
- Sending allowed only if `allow_send=True` **and** recipient in `send_allowlist` or allowed domains (`send_allow_domains`, default ru.naos.com/naos.com).
- Overdue reminders respect quiet hours and reminder intervals.
- Excel password can be overridden via `NAOS_EXCEL_PASSWORD` env var.
- Logs avoid bodies; UI shows SEND badge (red when live).

## Outlook/Excel notes

- Classic Outlook COM only; New Outlook is blocked with a clear error.
- Restrict filters use US 12-hour format for reliability.
- Sender SMTP resolved via ExchangeUser.PrimarySmtpAddress -> Address -> SenderEmailAddress fallback.
- Text cleaning strips quotes/signatures before hashing/storing.
- Excel sheets: Zayavki / Prosrochki / KPI / Konflikty / Statusy & hints. Editable columns: Status, Responsible, Comment, Priority.

## Tests

- Run all: `pytest`
- Coverage suggestion: `pytest --cov=core --cov=ui` (pytest-cov can be added if needed).
- QA driver: `python cli.py qa-full` (runs pytest + semi-E2E driver in safe/preview).

## Release

- Build PyInstaller: `build.bat` (creates `dist/naos_sla`).
- For delivery: sign executables if possible to reduce SmartScreen friction.

## Data locations

- Default data dir: `%APPDATA%/NAOS_SLA_TRACKER` (db, logs, tickets.xlsx, config.json).
- Config fields: sender filter (mode/value), mailbox/folder, safe_mode/allow_send, reminder/quiet hours, sla_by_priority, escalation_matrix, docs_url/sharepoint_url.

## Credits

- Developer: Tyoma (ShapArt). Please keep SAFE unless you intentionally need SEND.
