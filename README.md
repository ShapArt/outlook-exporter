# Outlook SLA Toolkit

**Windows desktop tooling for turning Outlook requests into structured SLA state, Excel reports and reminders.**

I originally built this around a real support workflow where incoming requests lived in Outlook and their status was tracked manually. The project reads mail through Classic Outlook COM, normalises requests into SQLite, calculates SLA state, exports/synchronises Excel and can prepare or send deadline reminders under explicit safety settings.

## Flow

```text
Classic Outlook / MAPI
        │
        ▼
message ingest + sender/customer extraction
        │
        ▼
SQLite ticket state
        │
        ├──► SLA / status recalculation
        │
        ├──► Excel export ↔ reviewed Excel edits
        │
        └──► overdue plan → reminders

        CLI + PySide6 desktop UI
```

For the workflow it was built around, the automation removed roughly **an hour of repetitive work per day** across covered tasks and made overdue requests easier to see instead of leaving the state in mail threads and spreadsheets.

## What is actually implemented

### Outlook ingest

[`core/outlook.py`](core/outlook.py) connects to **Classic Outlook** through `win32com` / MAPI. It can:

- open the configured mailbox/folder;
- filter messages by sender rules;
- resolve SMTP addresses from Exchange/Outlook objects;
- extract an external customer address from sender, forwarded headers or reply recipients;
- detect whether Classic Outlook COM is available and report New Outlook when COM cannot be used.

This is intentionally a Windows-first tool. New Outlook does not expose the same COM path, so the code fails with a clear diagnostic instead of pretending the integration is portable.

### SLA state

[`core/sla.py`](core/sla.py) turns message history into ticket state. The model includes statuses such as:

`new · assigned · responded · resolved · waiting_customer · table · overdue`

It also calculates business-hour SLA time, processes responses and builds an overdue reminder plan.

### SQLite

[`core/db.py`](core/db.py) stores the operational state separately from Outlook. That makes the mailbox a source of events rather than the only place where the current status exists.

### Excel round-trip

[`core/excel.py`](core/excel.py) exports the current state to Excel and supports synchronising reviewed spreadsheet changes back into the local data model. Tests cover export behaviour, protection and Excel round-trip cases.

### Notifications

[`core/notify.py`](core/notify.py) handles reminder messages. Sending is gated by configuration: the CLI exposes safe mode / allow-send controls and reports the resulting plan before sending overdue notifications.

### Desktop UI and CLI

The same core logic is available through:

- [`ui/app.py`](ui/app.py) — PySide6 desktop interface;
- [`cli.py`](cli.py) — operational and diagnostic commands;
- [`launch_ui.py`](launch_ui.py) — UI launcher.

## Useful CLI commands

The CLI includes commands for the main workflow rather than only one export script:

```text
ingest             read recent Outlook messages
recalc             recalculate open ticket state
export             write the Excel view
process-responses  update state from mail replies
send-overdue       build/send overdue reminders
diagnose           check Outlook COM, paths and filters
qa-full / test-all run automated and semi-E2E checks
```

Run `python cli.py --help` for the exact arguments available in the current version.

## Run locally

Requirements:

- Windows;
- **Classic Outlook** configured for the mailbox you want to read;
- Python 3;
- access rights to the mailbox/folder in Outlook.

Install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Copy [`config.example.json`](config.example.json) to your local configuration and adjust mailbox/folder, SLA and safety settings.

Start the desktop UI:

```powershell
python launch_ui.py
```

Or diagnose the environment first:

```powershell
python cli.py diagnose
```

## Stack

`Python` · `pywin32` · `SQLite` · `pandas` · `openpyxl` · `PySide6` · `pytest`

The repository also contains packaging helpers for Windows/PyInstaller and QA scripts for semi-E2E checks.

## Repository map

```text
core/outlook.py   Outlook COM / MAPI integration
core/sla.py       ticket states, SLA and response processing
core/db.py        SQLite persistence
core/excel.py     Excel export / sync
core/notify.py    reminders
ui/app.py         PySide6 desktop UI
cli.py            operational CLI / diagnostics
outlook_extract.py legacy / direct extraction path
tests/            unit and workflow tests
qa/               QA runbooks and drivers
```

## Safety / operational constraints

- Mail sending can be disabled independently from reading/processing.
- `safe_mode` and allowlists exist for test/QA runs.
- Outlook availability is diagnosed explicitly before COM-dependent work.
- Sender filtering is configurable and diagnostics show when a filter removes the whole message window.
- Excel is treated as an operator surface around structured state, not as the only database.

## Tests

The repository has tests for SLA calculations, business hours, sender filters, customer-email extraction, DB constraints, status mapping and Excel round-trip behaviour.

```powershell
python -m pytest
```

There is also a QA command that combines pytest with a semi-E2E driver and writes `QA_REPORT.md`.

## По-русски

Это Windows-инструмент для рабочего процесса вокруг Outlook: забрать обращения из почты, привести их к нормальным статусам, хранить состояние в SQLite, считать SLA, выгружать/синхронизировать Excel и напоминать о просрочках.

Проект появился не как учебный “экспорт почты”, а из ежедневной рутины в NAOS. Самая полезная часть — связка `Outlook → структурированное состояние → SLA → Excel/напоминания`, которая убрала повторяющиеся ручные операции.

## License

See [LICENSE](LICENSE).
