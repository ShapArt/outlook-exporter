from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Sequence

from config import AppConfig
from core.logger import get_logger

log = get_logger(__name__)


@dataclass
class TicketRecord:
    id: Optional[int]
    conv_id: Optional[str]
    thread_key: str
    entry_id: Optional[str]
    first_received_utc: datetime
    sender: str
    subject: str
    body: str
    first_forward_utc: Optional[datetime]
    first_forward_to: Optional[str]
    first_reply_utc: Optional[datetime]
    first_reply_body: Optional[str]
    responsible: Optional[str]
    status: str
    last_status_utc: datetime
    days_without_update: int
    overdue: bool
    not_interesting: bool
    customer_email: Optional[str]
    is_repeat: bool
    repeat_hint: Optional[str]
    recommended_answer: Optional[str]
    match_score: Optional[float]
    topic: Optional[str]
    last_reminder_utc: Optional[datetime]
    priority: Optional[str]
    stable_id: Optional[str]
    last_updated_at: Optional[datetime]
    last_updated_by: Optional[str]
    data_source: Optional[str]
    comment: Optional[str]
    row_version: int = 1


def connect(db_path: Path, wal_mode: bool = True) -> sqlite3.Connection:
    conn = sqlite3.connect(
        db_path, timeout=10, isolation_level=None, detect_types=sqlite3.PARSE_DECLTYPES
    )
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON;")
    if wal_mode:
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
        except sqlite3.DatabaseError as exc:
            log.warning("WAL mode not available: %s", exc)
    return conn


def ensure_schema(cfg: AppConfig) -> None:
    cfg.paths.ensure()
    conn = connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS tickets (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          stable_id TEXT,
          conv_id TEXT,
          thread_key TEXT NOT NULL,
          entry_id TEXT,
          first_received_utc TEXT NOT NULL,
          sender TEXT NOT NULL,
          subject TEXT NOT NULL,
          body TEXT,
          first_forward_utc TEXT,
          first_forward_to TEXT,
          first_reply_utc TEXT,
          first_reply_body TEXT,
          responsible TEXT,
          status TEXT NOT NULL,
          last_status_utc TEXT NOT NULL,
          days_without_update INTEGER NOT NULL DEFAULT 0,
          overdue INTEGER NOT NULL DEFAULT 0,
          not_interesting INTEGER NOT NULL DEFAULT 0,
          customer_email TEXT,
          is_repeat INTEGER NOT NULL DEFAULT 0,
          repeat_hint TEXT,
          recommended_answer TEXT,
          match_score REAL,
          topic TEXT,
          last_reminder_utc TEXT,
          priority TEXT,
          last_updated_at TEXT,
          last_updated_by TEXT,
          data_source TEXT,
          comment TEXT,
          row_version INTEGER NOT NULL DEFAULT 1,
          UNIQUE(conv_id, thread_key),
          UNIQUE(stable_id)
        );
        CREATE TABLE IF NOT EXISTS answers (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          ticket_id INTEGER,
          question TEXT NOT NULL,
          answer TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS events (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          ticket_id INTEGER NOT NULL,
          event_type TEXT NOT NULL,
          status_before TEXT,
          status_after TEXT,
          source TEXT,
          event_dt_utc TEXT,
          raw_response TEXT,
          item_entry_id TEXT,
          UNIQUE(item_entry_id)
        );
        """
    )
    _ensure_column(cur, "tickets", "row_version", "INTEGER NOT NULL DEFAULT 1")
    _ensure_column(cur, "tickets", "last_reminder_utc", "TEXT")
    _ensure_column(cur, "tickets", "priority", "TEXT")
    _ensure_column(cur, "tickets", "stable_id", "TEXT")
    _ensure_column(cur, "tickets", "last_updated_at", "TEXT")
    _ensure_column(cur, "tickets", "last_updated_by", "TEXT")
    _ensure_column(cur, "tickets", "data_source", "TEXT")
    _ensure_column(cur, "tickets", "comment", "TEXT")
    _migrate_statuses(cur)
    conn.commit()
    conn.close()


# Маппинг старых русских статусов на новые английские
STATUS_MIGRATION = {
    "???‘<ü": "new",
    "' ‘?ø+?‘'ç ‘? ó?>>ç?": "in_table",
    "' ‘'ø+>ñ‘Åç": "assigned",
    "ÿç‘?ç??": "responded",
    "-øó‘?‘<‘'?": "resolved",
    "?‘??‘?‘??‘Øç??": "overdue",
    "?ç ñ?‘'ç‘?ç‘???": "not_interesting",
}


def _migrate_statuses(cur: sqlite3.Cursor) -> None:
    cur.execute("PRAGMA table_info(tickets)")
    cols = {r[1] for r in cur.fetchall()}
    if "status" not in cols:
        return
    for old, new in STATUS_MIGRATION.items():
        cur.execute("UPDATE tickets SET status=? WHERE status=?", (new, old))
    cur.execute("UPDATE tickets SET status='table' WHERE status='in_table'")


def _ensure_column(cur: sqlite3.Cursor, table: str, name: str, ddl: str) -> None:
    cur.execute(f"PRAGMA table_info({table})")
    cols = {r[1] for r in cur.fetchall()}
    if name not in cols:
        log.info("Adding column %s.%s", table, name)
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def upsert_ticket(conn: sqlite3.Connection, ticket: TicketRecord) -> int:
    payload = ticket.__dict__.copy()
    payload["overdue"] = 1 if ticket.overdue else 0
    payload["not_interesting"] = 1 if ticket.not_interesting else 0
    payload["is_repeat"] = 1 if ticket.is_repeat else 0
    cols = [
        "conv_id",
        "thread_key",
        "entry_id",
        "first_received_utc",
        "sender",
        "subject",
        "body",
        "first_forward_utc",
        "first_forward_to",
        "first_reply_utc",
        "first_reply_body",
        "responsible",
        "status",
        "last_status_utc",
        "days_without_update",
        "overdue",
        "not_interesting",
        "customer_email",
        "is_repeat",
        "repeat_hint",
        "recommended_answer",
        "match_score",
        "topic",
        "last_reminder_utc",
        "priority",
        "stable_id",
        "last_updated_at",
        "last_updated_by",
        "data_source",
        "comment",
        "row_version",
    ]
    placeholders = ", ".join("?" for _ in cols)
    update_expr = ", ".join(
        f"{c}=excluded.{c}" for c in cols if c not in ("first_received_utc",)
    )
    cur = conn.cursor()
    cur.execute(
        f"""
        INSERT INTO tickets ({", ".join(cols)})
        VALUES ({placeholders})
        ON CONFLICT(conv_id, thread_key) DO UPDATE SET
          {update_expr},
          row_version = tickets.row_version + 1
        """,
        tuple(payload.get(c) for c in cols),
    )
    if ticket.id is None:
        ticket_id = cur.lastrowid
    else:
        cur.execute(
            "SELECT id FROM tickets WHERE conv_id=? AND thread_key=?",
            (ticket.conv_id, ticket.thread_key),
        )
        row = cur.fetchone()
        ticket_id = row["id"] if row else ticket.id
    return int(ticket_id)


def log_event(
    conn: sqlite3.Connection,
    ticket_id: int,
    event_type: str,
    status_before: Optional[str],
    status_after: Optional[str],
    source: str,
    raw_response: Optional[str] = None,
    item_entry_id: Optional[str] = None,
) -> None:
    conn.execute(
        """
        INSERT OR IGNORE INTO events (ticket_id, event_type, status_before, status_after, source, event_dt_utc, raw_response, item_entry_id)
        VALUES (?, ?, ?, ?, ?, datetime('now'), ?, ?)
        """,
        (
            ticket_id,
            event_type,
            status_before,
            status_after,
            source,
            raw_response,
            item_entry_id,
        ),
    )


def fetch_tickets(
    conn: sqlite3.Connection, where: str = "", params: Sequence = ()
) -> List[sqlite3.Row]:
    sql = """
    SELECT
      id, conv_id, thread_key, entry_id, first_received_utc, sender, subject, body,
      first_forward_utc, first_forward_to, first_reply_utc, first_reply_body,
      responsible, status, last_status_utc, days_without_update, overdue, not_interesting,
      customer_email, is_repeat, repeat_hint, recommended_answer, match_score, topic, last_reminder_utc,
      priority, stable_id, last_updated_at, last_updated_by, data_source, comment, row_version
    FROM tickets
    """
    if where:
        sql += f" WHERE {where}"
    sql += " ORDER BY datetime(first_received_utc) DESC"
    cur = conn.cursor()
    cur.execute(sql, params)
    return cur.fetchall()


def mark_row_version(conn: sqlite3.Connection, ticket_id: int) -> None:
    conn.execute(
        "UPDATE tickets SET row_version = row_version + 1, last_status_utc=datetime('now') WHERE id=?",
        (ticket_id,),
    )


def seed_test_ticket(conn: sqlite3.Connection, status: str = "overdue") -> int:
    now = datetime.utcnow()
    subject = f"[TEST][SLA] Synthetic {status} ticket"
    days = 5 if status == "overdue" else 0
    record = TicketRecord(
        id=None,
        conv_id=f"test-{status}-{int(now.timestamp())}",
        thread_key=f"test-thread-{status}-{int(now.timestamp())}",
        entry_id=None,
        first_received_utc=now - timedelta(days=days),
        sender="test@example.com",
        subject=subject,
        body="Synthetic test ticket to validate SLA tracker.",
        first_forward_utc=None,
        first_forward_to=None,
        first_reply_utc=None,
        first_reply_body=None,
        responsible="responsible@example.com" if status != "new" else None,
        status=status,
        last_status_utc=now - timedelta(days=days),
        days_without_update=days,
        overdue=status == "overdue",
        not_interesting=False,
        customer_email=None,
        is_repeat=False,
        repeat_hint=None,
        recommended_answer=None,
        match_score=None,
        topic=None,
        last_reminder_utc=None,
        priority="p3",
        stable_id=None,
        last_updated_at=now,
        last_updated_by="test-all",
        data_source="test",
        comment=None,
    )
    ticket_id = upsert_ticket(conn, record)
    log_event(conn, ticket_id, "seed_test_ticket", None, record.status, "test-all")
    conn.commit()
    return ticket_id
