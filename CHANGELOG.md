# Changelog

## 1.0.0

- Added modular architecture (config, db, outlook wrapper, SLA logic, Excel, notifications).
- Introduced CLI (`cli.py`) with `ingest`, `recalc-open`, `export-xlsx`, `send-overdue`, `diagnose`, `test-all`, `ui`.
- Implemented test mode: seeds overdue ticket and sends `[TEST][SLA]` email to allowlist.
- Added Excel export with hidden `ticket_id` and `row_version`, atomic writes, and optimistic-lock sync path.
- Added PySide6 UI with environment indicator, core actions, and log viewer.
- Added logging with rotation to `%APPDATA%/NAOS_SLA_TRACKER/logs`.
- Provided PyInstaller spec (`naos_sla.spec`) and build script (`build.bat`).
