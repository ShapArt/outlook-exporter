import tempfile

from config import AppConfig, Paths
from core import db, sla


def _temp_cfg():
    tmp = tempfile.mkdtemp()
    paths = Paths(
        appdata_dir=tmp,
        log_dir=f"{tmp}/logs",
        db_path=f"{tmp}/test.sqlite3",
        excel_path=f"{tmp}/tickets.xlsx",
        backup_dir=f"{tmp}/backups",
    )
    return AppConfig(paths=paths)


def test_overdue_plan_respects_interval_and_missing_owner():
    cfg = _temp_cfg()
    cfg.quiet_hours_start = 25
    cfg.quiet_hours_end = 0
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    with_owner = db.seed_test_ticket(conn, status=sla.STATUS_OVERDUE)
    no_owner = db.seed_test_ticket(conn, status=sla.STATUS_OVERDUE)
    conn.execute("UPDATE tickets SET responsible=NULL WHERE id=?", (no_owner,))
    conn.commit()
    conn.close()
    sla.mark_reminder_sent(cfg, with_owner)

    plan = sla.overdue_plan(cfg)
    ids_interval = {r["id"] for r in plan["skip_interval"]}
    ids_no_owner = {r["id"] for r in plan["skip_no_responsible"]}
    ids_send = {r["id"] for r in plan["send"]}

    assert with_owner not in ids_send
    assert no_owner in ids_no_owner
    assert all(r["id"] not in ids_interval for r in plan["send"])
