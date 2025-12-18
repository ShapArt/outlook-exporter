from __future__ import annotations

import re
import sqlite3
from datetime import datetime, time, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from config import AppConfig
from core import db, recommend
from core.db import TicketRecord
from core.logger import get_logger
from core.outlook import OutlookClient, extract_customer_email, get_sender_smtp
from core.utils import (
    clean_text,
    compute_stable_id,
    normalize_subject,
    passes_sender_filter,
    safe_get,
)

log = get_logger(__name__)

# Statuses are ASCII codes; RU labels are for display.
STATUS_NEW = "new"
STATUS_ASSIGNED = "assigned"
STATUS_RESPONDED = "responded"
STATUS_RESOLVED = "resolved"
STATUS_WAITING_CUSTOMER = "waiting_customer"
STATUS_TABLE = "table"
STATUS_OTIP = "otip"
STATUS_OVERDUE = "overdue"
STATUS_NOT_INTERESTING = "not_interesting"

# Historical alias kept for migration/compatibility
STATUS_IN_TABLE = "in_table"
STATUS_ALIASES = {
    STATUS_IN_TABLE: STATUS_TABLE,
}

STATUS_MODEL: List[Tuple[str, str]] = [
    (STATUS_NEW, "Новое"),
    (STATUS_ASSIGNED, "В работе"),
    (STATUS_RESPONDED, "Дан ответ"),
    (STATUS_RESOLVED, "Закрыто"),
    (STATUS_WAITING_CUSTOMER, "Ждем клиента"),
    (STATUS_TABLE, "Требует данных/таблица"),
    (STATUS_OTIP, "OTIP"),
    (STATUS_OVERDUE, "Просрочка SLA"),
    (STATUS_NOT_INTERESTING, "Неинтересно/спам"),
]

STATUS_LABELS_RU: Dict[str, str] = {code: label for code, label in STATUS_MODEL}

STATUS_ORDER = {
    STATUS_NEW: 0,
    STATUS_TABLE: 1,
    STATUS_IN_TABLE: 1,
    STATUS_ASSIGNED: 2,
    STATUS_RESPONDED: 3,
    STATUS_OVERDUE: 4,
    STATUS_RESOLVED: 5,
    STATUS_NOT_INTERESTING: 6,
    STATUS_WAITING_CUSTOMER: 2,
    STATUS_OTIP: 2,
}


def normalize_status_code(code: Optional[str]) -> Optional[str]:
    if code is None:
        return None
    code_l = str(code).strip().lower()
    return STATUS_ALIASES.get(code_l, code_l if code_l in STATUS_LABELS_RU else code_l)


def status_code_to_label(code: str) -> str:
    canonical = normalize_status_code(code)
    return STATUS_LABELS_RU.get(canonical, code or "")


def status_text_to_code(text: str) -> Optional[str]:
    if not text:
        return None
    t = str(text).strip().lower()
    code = STATUS_TEXT_MAP.get(t)
    if code:
        return code
    if t in STATUS_LABELS_RU:
        return t
    if t in STATUS_ALIASES:
        return STATUS_ALIASES[t]
    t_clean = re.sub(
        r"[^a-z\u0430-\u044f\u0451 0-9_]+", "", t, flags=re.IGNORECASE
    ).strip()
    return STATUS_TEXT_MAP.get(t_clean) or STATUS_TEXT_MAP.get(
        t_clean.replace(" ", "_")
    )


# Normalized lower-case inputs mapped to canonical codes (RU + EN + legacy)
STATUS_TEXT_MAP: Dict[str, str] = {
    "новое": STATUS_NEW,
    "new": STATUS_NEW,
    "vhod": STATUS_NEW,
    "вход": STATUS_NEW,
    "входящее": STATUS_NEW,
    "in_table": STATUS_TABLE,
    "table": STATUS_TABLE,
    "таблица": STATUS_TABLE,
    "требует таблицы": STATUS_TABLE,
    "требует данных": STATUS_TABLE,
    "требует данных/таблица": STATUS_TABLE,
    "назначено": STATUS_ASSIGNED,
    "в работе": STATUS_ASSIGNED,
    "работа": STATUS_ASSIGNED,
    "переслано": STATUS_ASSIGNED,
    "forwarded": STATUS_ASSIGNED,
    "forward": STATUS_ASSIGNED,
    "assigned": STATUS_ASSIGNED,
    "naznacheno": STATUS_ASSIGNED,
    "ответ": STATUS_RESPONDED,
    "дан ответ": STATUS_RESPONDED,
    "responded": STATUS_RESPONDED,
    "ok": STATUS_RESPONDED,
    "ок": STATUS_RESPONDED,
    "закрыто": STATUS_RESOLVED,
    "закрыть": STATUS_RESOLVED,
    "resolved": STATUS_RESOLVED,
    "done": STATUS_RESOLVED,
    "close": STATUS_RESOLVED,
    "closed": STATUS_RESOLVED,
    "неинтересно": STATUS_NOT_INTERESTING,
    "неинтересно/спам": STATUS_NOT_INTERESTING,
    "не наш": STATUS_NOT_INTERESTING,
    "спам": STATUS_NOT_INTERESTING,
    "not_interesting": STATUS_NOT_INTERESTING,
    "просрочка": STATUS_OVERDUE,
    "просрочка sla": STATUS_OVERDUE,
    "overdue": STATUS_OVERDUE,
    "waiting_customer": STATUS_WAITING_CUSTOMER,
    "waiting": STATUS_WAITING_CUSTOMER,
    "waiting customer": STATUS_WAITING_CUSTOMER,
    "need more time": STATUS_WAITING_CUSTOMER,
    "more time": STATUS_WAITING_CUSTOMER,
    "нужно время": STATUS_WAITING_CUSTOMER,
    "ждём клиента": STATUS_WAITING_CUSTOMER,
    "ждем клиента": STATUS_WAITING_CUSTOMER,
    "ожидаем клиента": STATUS_WAITING_CUSTOMER,
    "ждём": STATUS_WAITING_CUSTOMER,
    "ждем": STATUS_WAITING_CUSTOMER,
    "ждём ответ": STATUS_WAITING_CUSTOMER,
    "ждем ответ": STATUS_WAITING_CUSTOMER,
    "otip": STATUS_OTIP,
    "отип": STATUS_OTIP,
}


def derive_base_status(
    has_forward: bool, has_reply: bool, requires_table: bool, is_uninteresting: bool
) -> str:
    if is_uninteresting:
        return STATUS_NOT_INTERESTING
    if requires_table:
        return STATUS_TABLE
    if has_reply:
        return STATUS_RESPONDED
    if has_forward:
        return STATUS_ASSIGNED
    # Prototype rule: без ответа и пересылки считаем "В работе", чтобы не висело "Новое"
    return STATUS_ASSIGNED


def raise_status(current: str, candidate: str) -> str:
    if STATUS_ORDER.get(candidate, -1) > STATUS_ORDER.get(current, -1):
        return candidate
    return current


def business_hours_between(
    start: datetime,
    end: datetime,
    start_hour: int = 10,
    end_hour: int = 19,
    holidays: Optional[Iterable[str]] = None,
) -> float:
    """Посчитать рабочие часы (пн-пт, start_hour-end_hour), игнорируя выходные и список holidays (YYYY-MM-DD)."""
    if end <= start:
        return 0.0
    holidays_set = set(holidays or [])
    total = 0.0
    current = start
    end_date = end.date()
    while current.date() <= end_date:
        is_weekend = current.weekday() >= 5
        is_holiday = current.date().isoformat() in holidays_set
        if not is_weekend and not is_holiday:
            sh = max(0, min(23, start_hour))
            eh = end_hour
            if eh >= 24:
                day_end = datetime.combine(
                    current.date(), time(hour=23, minute=59, second=59)
                )
            else:
                day_end = datetime.combine(current.date(), time(hour=max(sh, eh)))
            day_start = datetime.combine(current.date(), time(hour=sh))
            interval_start = max(current, day_start)
            interval_end = min(end, day_end)
            if interval_end > interval_start:
                total += (interval_end - interval_start).total_seconds() / 3600.0
        current = datetime.combine(current.date(), time.min) + timedelta(days=1)
    return max(total, 0.0)


def ingest_range(cfg: AppConfig, days: int, outlook_factory=None) -> int:
    db.ensure_schema(cfg)
    start_dt = datetime.now() - timedelta(days=days)
    processed = 0
    skipped_sender_filter = 0
    skipped_non_mailitem = 0
    factory = outlook_factory or OutlookClient
    with factory(cfg) as outlook:
        sent_idx = _build_sent_index(outlook, start_dt)
        conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
        cur = conn.cursor()
        cur.execute("BEGIN")
        try:
            for msg in outlook.iter_messages(start_dt):
                cls = safe_get(msg, "Class", 0)
                if cls != 43:
                    skipped_non_mailitem += 1
                    continue
                received_raw = safe_get(msg, "ReceivedTime")
                if not received_raw:
                    continue
                # pywintypes datetime -> python datetime
                received = _to_naive(received_raw)
                if not isinstance(received, datetime):
                    continue
                sender = get_sender_smtp(msg)
                if not passes_sender_filter(sender, cfg):
                    skipped_sender_filter += 1
                    continue

                subject = safe_get(msg, "Subject", "") or ""
                body_raw = safe_get(msg, "Body", "") or ""
                body = clean_text(body_raw)
                conv_id = safe_get(msg, "ConversationID")
                norm_subj = normalize_subject(subject)
                thread_key = (conv_id or norm_subj).lower()
                entry_id = safe_get(msg, "EntryID")

                sent_info = sent_idx.get(thread_key, {"replies": [], "forwards": []})
                first_reply = _find_first_after(sent_info["replies"], received)
                first_forward = _find_first_after(sent_info["forwards"], received)

                requires_table = False
                is_uninteresting = False

                base_status = derive_base_status(
                    bool(first_forward[0]),
                    bool(first_reply[0]),
                    requires_table,
                    is_uninteresting,
                )
                last_status_utc = first_reply[0] or first_forward[0] or received
                stable_id = compute_stable_id(
                    conv_id, received, sender, norm_subj, body
                )
                rec_answer = ""
                customer_email = extract_customer_email(
                    msg,
                    sender,
                    body_raw or "",
                    subject or "",
                    cfg.customer_internal_domains,
                )
                record = TicketRecord(
                    id=None,
                    conv_id=conv_id,
                    thread_key=thread_key,
                    entry_id=entry_id,
                    first_received_utc=received,
                    sender=sender,
                    customer_email=customer_email,
                    subject=subject,
                    body=body,
                    first_forward_utc=first_forward[0],
                    first_forward_to=first_forward[1],
                    first_reply_utc=first_reply[0],
                    first_reply_body=first_reply[1],
                    responsible=first_forward[1],
                    status=base_status,
                    last_status_utc=last_status_utc,
                    days_without_update=max(
                        0, (datetime.utcnow().date() - last_status_utc.date()).days
                    ),
                    overdue=False,
                    not_interesting=is_uninteresting,
                    is_repeat=False,
                    repeat_hint=None,
                    recommended_answer=rec_answer,
                    match_score=None,
                    topic=None,
                    last_reminder_utc=None,
                    priority=None,
                    stable_id=stable_id,
                    last_updated_at=datetime.utcnow(),
                    last_updated_by="ingest",
                    data_source="outlook",
                    comment=None,
                )
                db.upsert_ticket(conn, record)
                processed += 1
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    log.info(
        "Ingest summary: processed=%s, skipped_sender_filter=%s, skipped_non_mailitem=%s",
        processed,
        skipped_sender_filter,
        skipped_non_mailitem,
    )
    return processed


def recalc_open(cfg: AppConfig) -> int:
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, status, last_status_utc, days_without_update, first_received_utc, first_reply_utc, priority
        FROM tickets
        WHERE status NOT IN (?, ?, ?, ?)
        """,
        (STATUS_RESOLVED, STATUS_NOT_INTERESTING, STATUS_TABLE, STATUS_IN_TABLE),
    )
    rows = cur.fetchall()
    updated = 0
    now = datetime.utcnow()
    holidays = {h for h in cfg.holidays}
    for r in rows:
        current_status = normalize_status_code(r["status"]) or r["status"]
        last_status = datetime.fromisoformat(r["last_status_utc"])
        days = max(0, (now.date() - last_status.date()).days)
        priority = r["priority"] or "p3"
        sla_cfg = cfg.sla_by_priority.get(priority, cfg.sla_by_priority.get("p3", {}))
        fhours = float(sla_cfg.get("first_response_hours", cfg.overdue_days * 24))
        rhours = float(sla_cfg.get("resolution_hours", cfg.overdue_days * 24))

        recvd = datetime.fromisoformat(r["first_received_utc"])
        bus_hours_since_recv = business_hours_between(
            recvd,
            now,
            cfg.business_hours_start,
            cfg.business_hours_end,
            holidays=holidays,
        )
        first_reply = r["first_reply_utc"]
        first_reply_dt = datetime.fromisoformat(first_reply) if first_reply else None

        if current_status == STATUS_WAITING_CUSTOMER:
            conn.execute(
                "UPDATE tickets SET days_without_update=?, overdue=0 WHERE id=?",
                (days, r["id"]),
            )
            updated += 1
            continue

        overdue = False
        if not first_reply_dt and bus_hours_since_recv >= fhours:
            overdue = True
        elif bus_hours_since_recv >= rhours:
            overdue = True

        status = current_status
        if overdue and status not in (
            STATUS_OVERDUE,
            STATUS_RESPONDED,
            STATUS_ASSIGNED,
        ):
            status = STATUS_OVERDUE
        if (
            status in (STATUS_RESPONDED, STATUS_ASSIGNED)
            and bus_hours_since_recv >= rhours
        ):
            status = STATUS_OVERDUE

        conn.execute(
            "UPDATE tickets SET days_without_update=?, overdue=?, status=?, last_status_utc=? WHERE id=?",
            (days, int(overdue), status, now.isoformat(), r["id"]),
        )
        updated += 1
    conn.commit()
    conn.close()
    try:
        recommend.update_recommendations(cfg)
    except Exception as exc:
        log.warning("Recommendations refresh failed: %s", exc)
    return updated


def map_status_text(text: str) -> Optional[str]:
    return status_text_to_code(text)


def map_voting_response(text: str) -> Optional[str]:
    if not text:
        return None
    return map_status_text(text)


def _parse_command_block(body: str) -> Dict[str, str]:
    commands: Dict[str, str] = {}
    if not body:
        return commands
    for line in body.splitlines():
        stripped = line.strip()
        if not stripped:
            break
        low = stripped.lower()
        if (
            stripped.startswith(">")
            or low.startswith("from:")
            or low.startswith("sent:")
            or low.startswith("от:")
            or low.startswith("дата:")
        ):
            break
        if not stripped.startswith("/"):
            continue
        parts = stripped[1:].split(maxsplit=1)
        if not parts:
            continue
        cmd = parts[0].lower()
        arg = parts[1].strip() if len(parts) > 1 else ""
        commands[cmd] = arg
    return commands


def _find_ticket_for_message(
    cur, conv_id: Optional[str], norm_subj: str
) -> Optional[sqlite3.Row]:
    if conv_id:
        cur.execute(
            "SELECT * FROM tickets WHERE conv_id=? ORDER BY datetime(first_received_utc) DESC LIMIT 1",
            (conv_id,),
        )
        row = cur.fetchone()
        if row:
            return row
    cur.execute(
        "SELECT * FROM tickets WHERE thread_key=? ORDER BY datetime(first_received_utc) DESC LIMIT 1",
        (norm_subj,),
    )
    return cur.fetchone()


def _send_confirmation(
    outlook: OutlookClient, msg, cfg: AppConfig, updates: Dict[str, str]
) -> None:
    if not updates:
        return
    try:
        reply = msg.Reply()
        summary = "; ".join(f"{k}={v}" for k, v in updates.items())
        reply.Body = f"Принято. Обновлено: {summary}\\n\\n" + (
            safe_get(msg, "Body", "") or ""
        )
        if cfg.safe_mode or not cfg.allow_send:
            reply.Display()
        else:
            reply.Send()
    except Exception as exc:
        log.warning("Failed to send confirmation: %s", exc)


def is_quiet_hours(cfg: AppConfig, now: Optional[datetime] = None) -> bool:
    now = now or datetime.now()
    start = cfg.quiet_hours_start
    end = cfg.quiet_hours_end
    hour = now.hour
    if start < end:
        return start <= hour < end
    return hour >= start or hour < end


def process_responses(cfg: AppConfig, days: int = 7, outlook_factory=None) -> int:
    """Обработать ответы: VotingResponse или команды /status /prio /owner /comment."""
    db.ensure_schema(cfg)
    updated = 0
    factory = outlook_factory or OutlookClient
    with factory(cfg) as outlook:
        start = datetime.now() - timedelta(days=days)
        conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
        cur = conn.cursor()
        for msg in outlook.iter_messages(start):
            if safe_get(msg, "Class", 0) != 43:
                continue
            subj = safe_get(msg, "Subject", "") or ""
            body = safe_get(msg, "Body", "") or ""
            vr = safe_get(msg, "VotingResponse", "") or ""
            conv_id = safe_get(msg, "ConversationID")
            entry_id = safe_get(msg, "EntryID")
            norm_subj = normalize_subject(subj).lower()
            commands = _parse_command_block(body)

            status = map_voting_response(vr)
            if not status and "status" in commands:
                status = map_status_text(commands.get("status"))
            if not status:
                m = re.search(
                    r"(status|\u0441\u0442\u0430\u0442\u0443\u0441)\s*[:=]\s*([\w\s-]+)",
                    subj + "\\n" + body,
                    re.IGNORECASE,
                )
                if m:
                    status = map_status_text(m.group(2))

            priority = commands.get("prio") or commands.get("priority")
            owner = commands.get("owner") or commands.get("responsible")
            comment = commands.get("comment")

            row = _find_ticket_for_message(cur, conv_id, norm_subj)
            if not row:
                continue

            updates: Dict[str, str] = {}
            if status:
                updates["status"] = status
            if owner:
                updates["responsible"] = owner
            if priority:
                updates["priority"] = priority

            if not updates:
                continue

            set_parts = []
            params = []
            touch_ts = False
            if "status" in updates:
                set_parts.append("status=?")
                params.append(updates["status"])
                set_parts.append("days_without_update=0")
                set_parts.append("overdue=0")
                touch_ts = True
            if "responsible" in updates:
                set_parts.append("responsible=?")
                params.append(updates["responsible"])
                touch_ts = True
            if "priority" in updates:
                set_parts.append("priority=?")
                params.append(updates["priority"])
                touch_ts = True
            if touch_ts:
                set_parts.append("last_status_utc=datetime('now')")
            set_parts.append("row_version=row_version+1")
            params.append(row["id"])
            cur.execute(f"UPDATE tickets SET {', '.join(set_parts)} WHERE id=?", params)

            db.log_event(
                conn,
                row["id"],
                "mail_response",
                row["status"],
                updates.get("status", row["status"]),
                "mail",
                raw_response=vr or subj,
                item_entry_id=entry_id,
            )
            if comment:
                db.log_event(
                    conn,
                    row["id"],
                    "comment",
                    row["status"],
                    updates.get("status", row["status"]),
                    "mail",
                    raw_response=comment,
                    item_entry_id=entry_id,
                )
            _send_confirmation(outlook, msg, cfg, updates)
            updated += 1
        conn.commit()
        conn.close()
    return updated


def overdue_plan(cfg: AppConfig) -> Dict[str, List[dict]]:
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT * FROM tickets
        WHERE status IN (?, ?)
        """,
        (STATUS_OVERDUE, STATUS_RESPONDED),
    )
    rows = cur.fetchall()
    conn.close()
    plan: Dict[str, List[dict]] = {
        "send": [],
        "skip_interval": [],
        "skip_quiet": [],
        "skip_no_responsible": [],
    }
    quiet = is_quiet_hours(cfg)
    for row in rows:
        row_dict = dict(row)
        if not row_dict.get("responsible"):
            row_dict["skip_reason"] = "no_responsible"
            plan["skip_no_responsible"].append(row_dict)
            continue
        last_reminder = row["last_reminder_utc"]
        if last_reminder:
            try:
                dt = datetime.fromisoformat(last_reminder)
                if (datetime.utcnow() - dt) < timedelta(
                    hours=cfg.reminder_interval_hours
                ):
                    row_dict["skip_reason"] = "interval"
                    plan["skip_interval"].append(row_dict)
                    continue
            except Exception:
                pass
        if quiet:
            row_dict["skip_reason"] = "quiet_hours"
            plan["skip_quiet"].append(row_dict)
            continue
        plan["send"].append(row_dict)
    return plan


def overdue_candidates(cfg: AppConfig) -> List[dict]:
    return overdue_plan(cfg)["send"]


def mark_reminder_sent(cfg: AppConfig, ticket_id: int) -> None:
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    conn.execute(
        "UPDATE tickets SET last_reminder_utc=datetime('now'), row_version=row_version+1 WHERE id=?",
        (ticket_id,),
    )
    conn.commit()
    conn.close()


def _build_sent_index(
    outlook: OutlookClient, since_dt: datetime
) -> Dict[str, Dict[str, List]]:
    idx: Dict[str, Dict[str, List]] = {}
    if not hasattr(outlook, "get_sent_folder"):
        return idx
    sent = outlook.get_sent_folder()
    items = sent.Items
    items.Sort("[SentOn]", True)
    restricted = items.Restrict(
        f"[SentOn] >= '{since_dt.strftime('%m/%d/%Y %I:%M %p')}'"
    )
    getter_first = safe_get(restricted, "GetFirst")
    item = getter_first() if getter_first else None
    while item:
        cls = safe_get(item, "Class", 0)
        if cls == 43:
            subject = safe_get(item, "Subject", "") or ""
            sent_on = _to_naive(safe_get(item, "SentOn"))
            conv_id = safe_get(item, "ConversationID")
            norm_subj = normalize_subject(subject)
            key = (conv_id or norm_subj).lower()
            rec = idx.setdefault(key, {"replies": [], "forwards": []})
            body = clean_text(safe_get(item, "Body", "") or "")
            subj_l = subject.lower()
            if subj_l.startswith("re:"):
                rec["replies"].append((sent_on, body))
            elif subj_l.startswith(("fw:", "fwd:")):
                to_val = safe_get(item, "To", "") or ""
                rec["forwards"].append((sent_on, to_val.split(";")[0].strip()))
        getter_next = safe_get(restricted, "GetNext")
        item = getter_next() if getter_next else None
    for rec in idx.values():
        rec["replies"].sort(key=lambda x: x[0])
        rec["forwards"].sort(key=lambda x: x[0])
    return idx


def _find_first_after(
    events: List[Tuple[datetime, str]], after_dt: datetime
) -> Tuple[Optional[datetime], Optional[str]]:
    after = _to_naive(after_dt)
    for dt, payload in events:
        ndt = _to_naive(dt)
        if ndt and after and ndt >= after:
            return ndt, payload
    return None, None


def _to_naive(dt_obj):
    if dt_obj is None:
        return None
    try:
        if hasattr(dt_obj, "timestamp"):
            dt_val = datetime.fromtimestamp(dt_obj.timestamp())
        else:
            dt_val = dt_obj
        if isinstance(dt_val, datetime) and dt_val.tzinfo is not None:
            return dt_val.replace(tzinfo=None)
        return dt_val
    except Exception:
        return None
