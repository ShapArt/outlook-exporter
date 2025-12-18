import tempfile
from datetime import datetime

from config import AppConfig, Paths
from core import db, sla
from core.outlook_iface import FakeMail, FakeOutlookClient


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


def test_ingest_range_with_fake_outlook():
    cfg = _cfg()
    msg = FakeMail(
        Subject="Test subject",
        Body="Body text",
        ReceivedTime=datetime.now(),
        EntryID="entry1",
        ConversationID="conv1",
        SenderEmailAddress="user@example.com",
    )
    client = FakeOutlookClient([msg])
    processed = sla.ingest_range(cfg, days=1, outlook_factory=lambda c: client)
    assert processed == 1
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row = conn.execute("SELECT * FROM tickets").fetchone()
    conn.close()
    assert row["subject"] == "Test subject"


def test_process_responses_fake_vote():
    cfg = _cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    tid = db.seed_test_ticket(conn, status=sla.STATUS_OVERDUE)
    row = conn.execute("SELECT conv_id FROM tickets WHERE id=?", (tid,)).fetchone()
    conn.commit()
    conn.close()
    msg = FakeMail(
        Subject="Re: test",
        Body="",
        ReceivedTime=datetime.now(),
        EntryID="resp1",
        ConversationID=row["conv_id"],
        VotingResponse="Закрыть",
    )
    client = FakeOutlookClient([msg])
    updated = sla.process_responses(cfg, days=7, outlook_factory=lambda c: client)
    assert updated >= 1
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row = conn.execute("SELECT status FROM tickets WHERE id=?", (tid,)).fetchone()
    conn.close()
    assert row["status"] == sla.STATUS_RESOLVED
