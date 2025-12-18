import tempfile
from datetime import datetime

from config import AppConfig, Paths
from core import db
from core.db import TicketRecord


def _cfg():
    tmp = tempfile.mkdtemp()
    return AppConfig(
        paths=Paths(
            appdata_dir=tmp,
            log_dir=f"{tmp}/logs",
            db_path=f"{tmp}/db.sqlite3",
            excel_path=f"{tmp}/t.xlsx",
            backup_dir=f"{tmp}/bak",
        )
    )


def test_unique_stable_id():
    cfg = _cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    now = datetime.utcnow()
    rec = TicketRecord(
        id=None,
        conv_id="c1",
        thread_key="t1",
        entry_id="e1",
        first_received_utc=now,
        sender="a@example.com",
        subject="s1",
        body="body",
        first_forward_utc=None,
        first_forward_to=None,
        first_reply_utc=None,
        first_reply_body=None,
        responsible=None,
        status="new",
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
        priority=None,
        stable_id="stable-1",
        last_updated_at=now,
        last_updated_by="test",
        data_source="test",
        comment=None,
    )
    db.upsert_ticket(conn, rec)
    rv_before = conn.execute(
        "SELECT row_version FROM tickets WHERE conv_id='c1'"
    ).fetchone()["row_version"]
    db.upsert_ticket(conn, rec)
    rv_after = conn.execute(
        "SELECT row_version FROM tickets WHERE conv_id='c1'"
    ).fetchone()["row_version"]
    assert rv_after > rv_before
    conn.close()


def test_event_idempotency():
    cfg = _cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    tid = db.seed_test_ticket(conn, status="new")
    db.log_event(
        conn, tid, "mail_response", None, "resolved", "mail", item_entry_id="x1"
    )
    db.log_event(
        conn, tid, "mail_response", None, "resolved", "mail", item_entry_id="x1"
    )
    events = conn.execute(
        "SELECT count(*) as c FROM events WHERE item_entry_id='x1'"
    ).fetchone()["c"]
    conn.close()
    assert events == 1
