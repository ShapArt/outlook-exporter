from __future__ import annotations

import argparse
import sys
from datetime import datetime

from config import AppConfig, default_config
from core import excel, notify, sla
from core.logger import setup_logging
from core.outlook import detect_outlook_environment


def cmd_ingest(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    days = args.days or cfg.ingest_days
    if args.ignore_filter:
        cfg.sender_filter = ""
    processed = sla.ingest_range(cfg, days)
    logger.info("Ingested %s messages (last %s days)", processed, days)


def cmd_recalc(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    updated = sla.recalc_open(cfg)
    logger.info("Recalc completed: %s tickets touched", updated)


def cmd_export(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    excel.export_excel(cfg, today_only=args.today_only)
    logger.info("Export finished")


def cmd_send_overdue(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    sla.recalc_open(cfg)
    plan = sla.overdue_plan(cfg)
    logger.info(
        "Overdue candidates: send=%s, interval_skip=%s, quiet_skip=%s, no_owner=%s",
        len(plan["send"]),
        len(plan["skip_interval"]),
        len(plan["skip_quiet"]),
        len(plan["skip_no_responsible"]),
    )
    for row in plan["send"]:
        entry_id = row.get("entry_id")
        notify.send_overdue_mail(
            cfg,
            row["responsible"],
            row["subject"],
            row["id"],
            priority=row.get("priority"),
            original_entry_id=entry_id,
        )
        sla.mark_reminder_sent(cfg, row["id"])


def cmd_diagnose(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    env = detect_outlook_environment()
    logger.info(
        "Outlook environment: classic=%s, new_outlook=%s, details=%s",
        env.classic_available,
        env.new_outlook_detected,
        env.details,
    )
    from core.outlook import OutlookClient

    with OutlookClient(cfg) as outlook:
        before, after, top10 = outlook.diagnose(args.days or cfg.ingest_days)
    print("Data dir:", cfg.paths.appdata_dir)
    print("Excel path:", cfg.paths.excel_path, "exists=", cfg.paths.excel_path.exists())
    print("DB path:", cfg.paths.db_path, "exists=", cfg.paths.db_path.exists())
    print("Mailbox:", cfg.mailbox or "<default profile>")
    print("Folder:", cfg.folder or "Inbox")
    print(
        f"Safe mode: {cfg.safe_mode}, allow_send: {cfg.allow_send}, allowlist: {', '.join(cfg.test_allowlist)}"
    )
    print(
        f"Sender filter mode: {cfg.sender_filter_mode}, value: {cfg.sender_filter_value or cfg.sender_filter}"
    )
    print(f"Messages in window: {before}, after sender filter: {after}")
    if before and after == 0:
        print(
            "Warning: sender filter might be too strict — попробуйте sender_filter_mode=off или domain."
        )
    print("Top real senders:")
    for sender, count in top10:
        print(f"  {sender or '<unknown>'}: {count}")
    try:
        probe = cfg.paths.appdata_dir / "_diag.tmp"
        probe.write_text("ok", encoding="utf-8")
        print("Write test: OK (data dir writable)")
        probe.unlink(missing_ok=True)
    except Exception as exc:
        print(f"Write test failed: {exc}")


def cmd_process_responses(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    updated = sla.process_responses(cfg, days=args.days or cfg.ingest_days)
    logger.info("Processed mail responses: %s", updated)
    print(f"Updated statuses from mail responses: {updated}")


def cmd_qa_full(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    report_lines = []
    env = detect_outlook_environment()
    report_lines.append(f"# QA Report {datetime.now().isoformat()}")
    report_lines.append(f"- Python: {sys.version}")
    report_lines.append(
        f"- Outlook: classic_available={env.classic_available}, new_outlook={env.new_outlook_detected}, details={env.details}"
    )
    report_lines.append(
        f"- Config: safe_mode={cfg.safe_mode}, allow_send={cfg.allow_send}, mailbox={cfg.mailbox}, folder={cfg.folder}"
    )
    # Step A: pytest
    import subprocess

    logger.info("Running pytest ...")
    test_proc = subprocess.run(
        [sys.executable, "-m", "pytest"], capture_output=True, text=True
    )
    report_lines.append("## Pytest")
    report_lines.append(f"exit_code={test_proc.returncode}")
    report_lines.append("```")
    report_lines.append(test_proc.stdout[-2000:])
    report_lines.append("```")
    if test_proc.returncode != 0:
        logger.error("Pytest failed; see QA_REPORT.md")
    # Step B: semi-e2e driver (safe mode by default, send if requested)
    logger.info("Running semi-e2e driver ...")
    driver_cmd = [sys.executable, "qa/tools/qa_e2e_driver.py"]
    if args.send:
        driver_cmd.append("--send")
    if getattr(args, "no_wait", False):
        driver_cmd.append("--no-wait")
    driver_proc = subprocess.run(driver_cmd, capture_output=True, text=True)
    report_lines.append("## Semi E2E (qa_e2e_driver)")
    report_lines.append(f"exit_code={driver_proc.returncode}")
    report_lines.append("```")
    report_lines.append(driver_proc.stdout[-2000:])
    report_lines.append("```")
    if driver_proc.returncode != 0:
        logger.error("qa_e2e_driver failed; see QA_REPORT.md")
    # Save report
    from pathlib import Path

    out_path = Path.cwd() / "QA_REPORT.md"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(report_lines), encoding="utf-8")
    logger.info("QA report saved to %s", out_path)
    print(f"QA report: {out_path}")


def cmd_test_all(args):
    cfg = load_cfg(args)
    if args.send:
        cfg.allow_send = True
        cfg.safe_mode = False
    if args.safe:
        cfg.safe_mode = True
    setup_logging(cfg.paths.log_dir)
    from core import db

    steps = []
    ok = steps.append
    fail_msgs = []

    try:
        db.ensure_schema(cfg)
        ok("DB schema ok")
    except Exception as exc:
        fail_msgs.append(f"DB schema FAILED: {exc}")

    try:
        conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
        ticket_ids = []
        ticket_ids.append(db.seed_test_ticket(conn, status="new"))
        ticket_ids.append(db.seed_test_ticket(conn, status="assigned"))
        ticket_ids.append(db.seed_test_ticket(conn, status="overdue"))
        conn.close()
        ok(f"Seeded tickets {ticket_ids}")
    except Exception as exc:
        fail_msgs.append(f"Seed FAILED: {exc}")

    try:
        updated = sla.recalc_open(cfg)
        ok(f"Recalc done ({updated} tickets)")
    except Exception as exc:
        fail_msgs.append(f"Recalc FAILED: {exc}")

    try:
        excel_path = excel.export_excel(cfg, today_only=False)
        ok(f"Excel exported: {excel_path}")
    except Exception as exc:
        fail_msgs.append(f"Excel export FAILED: {exc}")

    try:
        # emulate Excel edit: mark first row resolved if file exists
        from pathlib import Path

        import pandas as pd

        path = Path(cfg.paths.excel_path)
        if path.exists():
            df = pd.read_excel(path, sheet_name=excel.SHEET_TICKETS)
            if not df.empty and "Статус" in df.columns:
                df.loc[0, "Статус"] = sla.status_code_to_label(sla.STATUS_RESOLVED)
                df.to_excel(path, sheet_name=excel.SHEET_TICKETS, index=False)
                res = excel.sync_from_excel(cfg)
                ok(f"Excel->DB sync: {res}")
                # verify DB updated
                conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
                status = conn.execute(
                    "SELECT status FROM tickets WHERE id=?",
                    (int(df.loc[0, "ticket_id"]),),
                ).fetchone()
                conn.close()
                if status and status["status"] == sla.STATUS_RESOLVED:
                    ok("DB status updated to resolved after sync")
                else:
                    fail_msgs.append("DB status did not update after Excel sync")
    except Exception as exc:
        fail_msgs.append(f"Excel sync FAILED: {exc}")

    try:
        notify.send_test_mail(
            cfg,
            "[TEST][SLA] Overdue reminder (test-all)",
            "Synthetic reminder from test mode.",
        )
        ok("Test mail queued")
    except Exception as exc:
        fail_msgs.append(f"Test mail FAILED: {exc}")

    print("==== TEST MODE REPORT ====")
    for s in steps:
        print("[OK]", s)
    if fail_msgs:
        for f in fail_msgs:
            print("[FAIL]", f)
    if not fail_msgs:
        print("All steps PASS")
    print("Test mail dispatched to allowlist:", ", ".join(cfg.test_allowlist))


def cmd_sync_excel(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    res = excel.sync_from_excel(cfg)
    logger.info("Excel sync: %s", res)


def cmd_sync_all(args):
    cfg = load_cfg(args)
    logger = setup_logging(cfg.paths.log_dir)
    days = args.days or cfg.ingest_days
    logger.info("== Шаг A: Excel -> DB ==")
    sync_res = (
        excel.sync_from_excel(cfg)
        if not args.dry_run
        else {"updated": 0, "conflicts": 0, "missing": 0, "conflict_rows": []}
    )
    logger.info("Excel sync: %s", sync_res)
    logger.info("== Шаг B: Outlook ingest ==")
    processed = sla.ingest_range(cfg, days)
    logger.info("Ingested %s messages (last %s days)", processed, days)
    logger.info("== Шаг C: Recalc ==")
    updated = sla.recalc_open(cfg)
    logger.info("Recalc updated %s", updated)
    logger.info("== Шаг D: Export ==")
    excel.export_excel(
        cfg,
        conflicts=sync_res.get("conflict_rows") if isinstance(sync_res, dict) else None,
    )
    logger.info("== Шаг E: Notify preview ==")
    plan = sla.overdue_plan(cfg)
    for row in plan["send"]:
        logger.info(
            "Просрочка: #%s %s -> %s",
            row.get("id"),
            row.get("subject"),
            row.get("responsible"),
        )
    if plan["skip_interval"]:
        logger.info(
            "Пропущено из-за интервала: %s",
            [r.get("id") for r in plan["skip_interval"]],
        )
    if plan["skip_quiet"]:
        logger.info(
            "Пропущено (quiet hours): %s", [r.get("id") for r in plan["skip_quiet"]]
        )
    if plan["skip_no_responsible"]:
        logger.info(
            "Пропущено (нет ответственного): %s",
            [r.get("id") for r in plan["skip_no_responsible"]],
        )
    logger.info("Пайплайн sync-all завершён")


def cmd_ui(args):
    cfg = load_cfg(args)
    setup_logging(cfg.paths.log_dir)
    from ui.app import run_ui

    run_ui(cfg)


def load_cfg(args) -> AppConfig:
    if getattr(args, "config", None):
        return AppConfig.load(args.config)
    return default_config()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="naos_sla", description="NAOS SLA tracker (refactored)"
    )
    parser.add_argument("--config", type=str, help="Path to config.json")
    sub = parser.add_subparsers(dest="command", required=True)

    p_all = sub.add_parser(
        "sync-all",
        help="Полный цикл: Excel->DB, Outlook ingest, recalc, экспорт, notify (preview в safe)",
    )
    p_all.add_argument(
        "--days",
        type=int,
        default=None,
        help="Дней для ingest (если не задано — из конфига)",
    )
    p_all.add_argument(
        "--dry-run",
        action="store_true",
        help="Только показать, что будет обновлено из Excel",
    )
    p_all.set_defaults(func=cmd_sync_all)

    p_ingest = sub.add_parser(
        "ingest", help="Ingest messages for N days (default from config)"
    )
    p_ingest.add_argument("--days", type=int, default=None)
    p_ingest.add_argument(
        "--ignore-filter", action="store_true", help="Ignore sender_filter (ingest all)"
    )
    p_ingest.set_defaults(func=cmd_ingest)

    p_recalc = sub.add_parser("recalc-open", help="Recalculate open tickets")
    p_recalc.set_defaults(func=cmd_recalc)

    p_export = sub.add_parser("export-xlsx", help="Export tickets to Excel")
    p_export.add_argument("--today-only", action="store_true")
    p_export.set_defaults(func=cmd_export)

    p_sync = sub.add_parser("sync-excel", help="Sync changes from Excel back to DB")
    p_sync.set_defaults(func=cmd_sync_excel)

    p_send = sub.add_parser("send-overdue", help="Send overdue reminders")
    p_send.set_defaults(func=cmd_send_overdue)

    p_diag = sub.add_parser("diagnose", help="Diagnose mailbox usage and sender filter")
    p_diag.add_argument("--days", type=int, default=None)
    p_diag.set_defaults(func=cmd_diagnose)

    p_resp = sub.add_parser(
        "process-responses", help="Process voting/STATUS responses from mailbox"
    )
    p_resp.add_argument("--days", type=int, default=None)
    p_resp.set_defaults(func=cmd_process_responses)

    p_test = sub.add_parser("test-all", help="Seed test ticket and send test email")
    p_test.add_argument(
        "--send", action="store_true", help="Send email (disable safe/display)"
    )
    p_test.add_argument("--safe", action="store_true", help="Force safe/display mode")
    p_test.set_defaults(func=cmd_test_all)

    p_qa = sub.add_parser(
        "qa-full", help="Run full QA suite (pytest + semi-e2e + report)"
    )
    p_qa.add_argument(
        "--send",
        action="store_true",
        help="Allow sending mails in qa_e2e_driver (default display)",
    )
    p_qa.add_argument(
        "--no-wait", action="store_true", help="Skip manual Voting step (useful for CI)"
    )
    p_qa.set_defaults(func=cmd_qa_full)

    p_ui = sub.add_parser("ui", help="Start desktop UI")
    p_ui.set_defaults(func=cmd_ui)
    return parser


def main(argv: list[str] | None = None):
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main(sys.argv[1:])
