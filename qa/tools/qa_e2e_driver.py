from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config import AppConfig, default_config  # noqa: E402
from core import db, excel, notify, sla  # noqa: E402
from core.outlook import detect_outlook_environment  # noqa: E402


def main():
    parser = argparse.ArgumentParser(
        description="Semi E2E driver for QA demo (Artem overdue -> reminder -> voting)"
    )
    parser.add_argument("--config", type=str, help="Path to config.json")
    parser.add_argument(
        "--send", action="store_true", help="Allow sending instead of Display"
    )
    parser.add_argument(
        "--no-wait",
        action="store_true",
        help="Do not wait for manual voting (useful for CI)",
    )
    args = parser.parse_args()

    cfg = AppConfig.load(args.config) if args.config else default_config()
    if args.send:
        cfg.allow_send = True
        cfg.safe_mode = False
    cfg.paths.ensure()
    db.ensure_schema(cfg)

    env = detect_outlook_environment()
    print(
        f"Outlook env: classic={env.classic_available}, new={env.new_outlook_detected}, details={env.details}"
    )
    if not env.classic_available:
        print("WARNING: Classic Outlook COM unavailable. E2E steps will not execute.")

    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    # Step 1: seed or pick existing overdue ticket
    ticket_id = db.seed_test_ticket(conn, status=sla.STATUS_OVERDUE)
    # Force overdue by backdating
    force_days = cfg.qa_force_overdue_days or 5
    conn.execute(
        "UPDATE tickets SET first_received_utc=datetime('now', ?), last_status_utc=datetime('now', ?), responsible=?, entry_id=entry_id WHERE id=?",
        (
            f"-{force_days} day",
            f"-{force_days} day",
            cfg.qa_artem_email
            or (cfg.test_allowlist[0] if cfg.test_allowlist else None),
            ticket_id,
        ),
    )
    conn.commit()
    conn.close()
    print(
        f"Seeded overdue ticket #{ticket_id}, responsible={cfg.qa_artem_email or cfg.test_allowlist[0] if cfg.test_allowlist else None}"
    )

    # Step 2: recalc to ensure overdue flag
    sla.recalc_open(cfg)

    # Step 3: send overdue reminder (reply if entry_id present)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row = conn.execute("SELECT * FROM tickets WHERE id=?", (ticket_id,)).fetchone()
    conn.close()
    entry_id = row["entry_id"]
    notify.send_overdue_mail(
        cfg,
        row["responsible"]
        or (
            cfg.qa_artem_email or (cfg.test_allowlist[0] if cfg.test_allowlist else "")
        ),
        row["subject"],
        row["id"],
        priority=row["priority"],
        original_entry_id=entry_id,
        preview_only=not args.send,
    )
    print("Overdue reminder triggered (display/preview if safe_mode).")

    if not args.no_wait:
        print("\n=== Manual step ===")
        print(
            "Open Outlook, click Voting 'Closed' (or reply /status closed) and press Enter here when done..."
        )
        try:
            input()
        except KeyboardInterrupt:
            pass

    # Step 4: process responses
    updated = sla.process_responses(cfg, days=cfg.ingest_days or 7)
    print(f"Process responses updated: {updated}")

    # Step 5: export Excel and show status
    excel_path = excel.export_excel(cfg, today_only=False)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row_after = conn.execute(
        "SELECT status, overdue, row_version FROM tickets WHERE id=?", (ticket_id,)
    ).fetchone()
    conn.close()
    print(
        f"Ticket after responses: status={row_after['status']}, overdue={row_after['overdue']}, row_version={row_after['row_version']}"
    )
    print(f"Excel exported: {excel_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
