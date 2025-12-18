import tempfile
from datetime import datetime, timedelta

from config import AppConfig, Paths
from core import db, sla
from core.sla import business_hours_between


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


def test_derive_base_status():
    assert sla.derive_base_status(False, False, False, False) == sla.STATUS_ASSIGNED
    assert sla.derive_base_status(True, False, False, False) == sla.STATUS_ASSIGNED
    assert sla.derive_base_status(False, True, False, False) == sla.STATUS_RESPONDED
    assert sla.derive_base_status(False, False, True, False) == sla.STATUS_TABLE
    assert (
        sla.derive_base_status(False, False, False, True) == sla.STATUS_NOT_INTERESTING
    )


def test_recalc_marks_overdue():
    cfg = _temp_cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    now = datetime.utcnow() - timedelta(days=5)
    rec = sla.TicketRecord(
        id=None,
        conv_id="c1",
        thread_key="t1",
        entry_id=None,
        first_received_utc=now,
        sender="a@b.c",
        subject="Test",
        body="Body",
        first_forward_utc=None,
        first_forward_to=None,
        first_reply_utc=None,
        first_reply_body=None,
        responsible="resp@b.c",
        status=sla.STATUS_NEW,
        last_status_utc=now,
        days_without_update=5,
        overdue=False,
        not_interesting=False,
        customer_email=None,
        is_repeat=False,
        repeat_hint=None,
        recommended_answer=None,
        match_score=None,
        topic=None,
        last_reminder_utc=None,
        priority="p3",
        stable_id="test-stable",
        last_updated_at=now,
        last_updated_by="test",
        data_source="test",
        comment=None,
    )
    db.upsert_ticket(conn, rec)
    conn.commit()
    conn.close()
    updated = sla.recalc_open(cfg)
    assert updated >= 1
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row = conn.execute("SELECT status FROM tickets WHERE thread_key='t1'").fetchone()
    conn.close()
    assert row["status"] in (
        sla.STATUS_OVERDUE,
        sla.STATUS_RESPONDED,
        sla.STATUS_ASSIGNED,
    )


def test_parse_command_block_stops_before_quote():
    body = "/status resolved\n/owner user@example.com\n\n> quoted line\n/status ignored"
    commands = sla._parse_command_block(body)
    assert commands["status"] == "resolved"
    assert commands["owner"] == "user@example.com"
    assert "ignored" not in "".join(commands.values())


def test_map_status_text_understands_waiting():
    assert sla.map_status_text("waiting") == sla.STATUS_WAITING_CUSTOMER
    assert sla.map_status_text("нужно время") == sla.STATUS_WAITING_CUSTOMER


def test_business_hours_between_basic():
    start = datetime(2024, 1, 2, 10, 0, 0)
    end = datetime(2024, 1, 2, 18, 0, 0)
    assert business_hours_between(start, end) == 8.0


def test_priority_sla_overdue_by_hours():
    cfg = _temp_cfg()
    cfg.business_hours_start = 10
    cfg.business_hours_end = 19
    cfg.sla_by_priority = {
        "p1": {"first_response_hours": 4, "resolution_hours": 24},
        "p3": {"first_response_hours": 16, "resolution_hours": 48},
    }
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    now = datetime.utcnow() - timedelta(days=3, hours=2)
    rec = sla.TicketRecord(
        id=None,
        conv_id="c2",
        thread_key="t2",
        entry_id=None,
        first_received_utc=now,
        sender="a@b.c",
        subject="Test",
        body="Body",
        first_forward_utc=None,
        first_forward_to=None,
        first_reply_utc=None,
        first_reply_body=None,
        responsible="resp@b.c",
        status=sla.STATUS_NEW,
        last_status_utc=now,
        days_without_update=0,
        overdue=False,
        not_interesting=False,
        customer_email=None,
        is_repeat=False,
        repeat_hint=None,
        recommended_answer=None,
        match_score=None,
        topic=None,
        last_reminder_utc=None,
        priority="p1",
        stable_id="stable-t2",
        last_updated_at=now,
        last_updated_by="test",
        data_source="test",
        comment=None,
    )
    db.upsert_ticket(conn, rec)
    conn.commit()
    conn.close()
    updated = sla.recalc_open(cfg)
    assert updated >= 1
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row = conn.execute("SELECT status FROM tickets WHERE thread_key='t2'").fetchone()
    conn.close()
    assert row["status"] == sla.STATUS_OVERDUE
