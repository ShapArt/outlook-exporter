from __future__ import annotations

import hashlib
import re
import time
from datetime import datetime
from typing import Any, Callable, Optional, Tuple

import pythoncom

from config import AppConfig


def dt_to_restrict_str(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")


def safe_get(obj: Any, attr: str, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def safe_call(
    fn: Callable[[], Any], default=None, retries: int = 2, delay: float = 0.05
):
    for attempt in range(retries + 1):
        try:
            return fn()
        except pythoncom.com_error:
            if attempt < retries:
                time.sleep(delay * (attempt + 1))
            else:
                return default
        except Exception:
            return default


def normalize_subject(subject: str) -> str:
    if not subject:
        return ""
    return re.sub(r"^((re:|fw:|fwd:)\s*)+", "", subject, flags=re.IGNORECASE).strip()


def clean_text(text: str) -> str:
    """Strip quotes/signatures and normalise line breaks from Outlook bodies."""
    if not text:
        return ""

    reply_markers = (
        "from:",
        "to:",
        "subject:",
        "sent:",
        "от:",
        "кому:",
        "тема:",
        "отправлено:",
        "ответ от",
        "reply-to:",
        "----original message----",
    )
    signature_markers = (
        "--",
        "__",
        "best regards",
        "kind regards",
        "regards",
        "cheers",
        "thanks",
        "thank you",
        "с уважением",
        "спасибо",
        "с наилучшими пожеланиями",
        "отправлено из",
    )

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in text.split("\n")]
    cleaned: list[str] = []
    for ln in lines:
        stripped = ln.strip()
        low = stripped.lower()
        # Skip quoted blocks
        if stripped.startswith(">"):
            continue
        if any(low.startswith(pfx) for pfx in reply_markers):
            continue
        if re.match(r"^[-_]{5,}\s*$", stripped):
            break
        if any(low.startswith(sig) for sig in signature_markers):
            break
        cleaned.append(stripped.rstrip())

    # Trim empty lines at edges
    while cleaned and not cleaned[0].strip():
        cleaned.pop(0)
    while cleaned and not cleaned[-1].strip():
        cleaned.pop()
    return "\n".join(cleaned)


def local_today_range() -> Tuple[datetime, datetime]:
    now = datetime.now()
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start.replace(hour=23, minute=59, second=59)
    return start, end


def compute_stable_id(
    conv_id: Optional[str], received: datetime, sender: str, subject: str, body: str
) -> str:
    parts = [
        conv_id or "",
        received.isoformat(),
        (sender or "").lower(),
        (subject or "").strip().lower(),
        (body or "")[:200],
    ]
    raw = "|".join(parts)
    return hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest()


def _normalize_sender_filter(cfg: AppConfig) -> tuple[str, str]:
    mode = (cfg.sender_filter_mode or "").lower().strip() or "off"
    value = (cfg.sender_filter_value or cfg.sender_filter or "").lower().strip()
    if mode == "domain" and value.startswith("@"):
        value = value[1:]
    return mode, value


def passes_sender_filter(sender: str, cfg: AppConfig) -> bool:
    """Apply sender filter modes: off|contains|equals|domain."""
    sender_l = (sender or "").lower()
    mode, value = _normalize_sender_filter(cfg)
    if mode == "off" or not value:
        return True
    if mode == "contains":
        return value in sender_l
    if mode == "equals":
        return sender_l == value
    if mode == "domain":
        return sender_l.endswith("@" + value) or sender_l.endswith("." + value)
    return True
