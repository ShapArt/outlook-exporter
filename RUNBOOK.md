# RUNBOOK - NAOS SLA Tracker

## Daily flow (morning/evening)

1. Open UI (`python cli.py ui`) or exe. Check SEND badge (should be SAFE unless approved) and Outlook status banner.
2. Run full pipeline (UI button “Full sync”) or CLI `python cli.py sync-all`:
   - Excel -> DB sync (conflicts logged as events `excel_conflict`).
   - Ingest Outlook for N days (sender filter applied).
   - Recalc SLA/overdue.
   - Export Excel (password-protected; editable: Status/Responsible/Comment/Priority).
   - Process responses (Voting + /status /prio /owner /comment).
   - Send overdue reminders (preview if SAFE/quiet hours/interval/no owner blocked).
3. Check metrics (KPI/overdue/conflicts) and conflicts list if any.
4. For single ticket: select row -> Details, update status/responsible/priority, then send/preview reminders or replies (respecting policies).

## SAFE vs SEND

- SAFE (default): mails are displayed only. SEND requires confirmation and allowlist/allowed domains.
- Allowed domains: `send_allow_domains` (config.json). Explicit allowlist: `send_allowlist`/`test_allowlist`.
- Quiet hours + `reminder_interval_hours` gate reminder sending.

## Outlook/Excel troubleshooting

- **New Outlook**: not supported. Switch to Classic Outlook; error banner will warn if COM unavailable.
- **Excel locked**: export writes `.bak` and `_pending.xlsx` if file is locked; close Excel and rerun.
- **No responses processed**: check sender filter, VotingResponse, subject normalization (/status etc.).
- **SMTP resolution**: uses ExchangeUser.PrimarySmtpAddress -> Address -> SenderEmailAddress.

## Config quick refs

- `%APPDATA%/NAOS_SLA_TRACKER/config.json` auto-generated on first run.
- Key fields: mailbox, folder, sender_filter_mode/value, safe_mode, allow_send, send_allow_domains, reminder/quiet hours, excel_password (or env `NAOS_EXCEL_PASSWORD`), docs_url/sharepoint_url.

## QA / semi-E2E

- `python cli.py qa-full` (pytest + `qa/tools/qa_e2e_driver.py`).
- Semi-E2E driver: seeds overdue ticket, sends reminder (preview if safe), waits for manual Voting, processes responses, exports Excel.

## Manual diagnostics

- `python cli.py diagnose --days N` prints sender filter stats, top senders, data paths, Outlook COM status.
- Logs: `%APPDATA%/NAOS_SLA_TRACKER/logs/outlook_sla.log`.

## Known safety rails

- New Outlook blocked with clear error.
- SEND requires confirmation; badge turns red in UI.
- Recipients filtered by allowlist/domains; blocked addresses logged.
- Logging avoids email bodies; UI help available via Help button.
