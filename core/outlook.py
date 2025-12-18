from __future__ import annotations

import re
import subprocess
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Iterable, List, Optional, Tuple

import pythoncom
import win32com.client

from config import AppConfig
from core.logger import get_logger
from core.outlook_iface import OutlookLike
from core.utils import dt_to_restrict_str, passes_sender_filter, safe_call, safe_get

log = get_logger(__name__)

# Fallback numeric folder IDs to avoid missing win32 constants
FOLDER_INBOX = 6
FOLDER_SENT = 5


def get_sender_smtp(mail) -> str:
    sender = safe_get(mail, "Sender")
    if sender:
        exch_user = safe_call(lambda: sender.GetExchangeUser(), None)
        if exch_user:
            smtp = safe_get(exch_user, "PrimarySmtpAddress")
            if smtp:
                return str(smtp).lower()
        addr = safe_get(sender, "Address")
        if addr:
            return str(addr).lower()
    addr2 = safe_get(mail, "SenderEmailAddress", "")
    return str(addr2 or "").lower()


def _is_internal(email: str, internal_domains: Optional[List[str]]) -> bool:
    if not email:
        return False
    domains = [d.lower().strip() for d in (internal_domains or []) if d]
    email_l = email.lower()
    for dom in domains:
        if email_l.endswith("@" + dom) or email_l.endswith("." + dom):
            return True
    return False


def extract_customer_email(
    msg,
    sender_smtp: str,
    body_text: str,
    subject: str,
    internal_domains: Optional[List[str]] = None,
) -> Optional[str]:
    """
    Derive customer email with priority:
    1) External sender (not internal) -> sender_smtp
    2) Forwarded markers in body (From/От/Reply-To/Email/Кому) -> first email
    3) Outlook reply-to/recipients fields -> first non-internal
    """
    sender_smtp = (sender_smtp or "").strip().lower()
    # 1) external sender
    if sender_smtp and not _is_internal(sender_smtp, internal_domains):
        return sender_smtp

    # 2) scan body top lines for forwarded headers
    if body_text:
        lines = body_text.splitlines()
        top_lines = lines[:80]
        email_re = re.compile(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)")
        for ln in top_lines:
            low = ln.lower()
            if any(
                low.startswith(prefix)
                for prefix in (
                    "from:",
                    "reply-to:",
                    "email:",
                    "e-mail:",
                    "от:",
                    "кому:",
                )
            ):
                m = email_re.search(ln)
                if m:
                    cand = m.group(1).lower()
                    if not _is_internal(cand, internal_domains):
                        return cand
        # fallback: first email anywhere in top lines
        for ln in top_lines:
            m = email_re.search(ln)
            if m:
                cand = m.group(1).lower()
                if not _is_internal(cand, internal_domains):
                    return cand

    # 3) ReplyRecipients / ReplyTo
    try:
        reply_to = safe_get(msg, "ReplyTo") or ""
        if reply_to:
            reply_to = str(reply_to).lower()
            if reply_to and not _is_internal(reply_to, internal_domains):
                return reply_to
    except Exception:
        pass
    recipients = safe_get(msg, "ReplyRecipients")
    if recipients:
        try:
            enum = getattr(recipients, "__iter__", None)
            if enum:
                for rec in recipients:
                    addr = safe_get(rec, "Address")
                    if addr:
                        addr_l = str(addr).lower()
                        if addr_l and not _is_internal(addr_l, internal_domains):
                            return addr_l
            else:
                # COM collection style
                count = safe_get(recipients, "Count", 0) or 0
                getter = safe_get(recipients, "Item")
                if getter:
                    for i in range(1, count + 1):
                        rec = getter(i)
                        addr = safe_get(rec, "Address")
                        addr_l = str(addr or "").lower()
                        if addr_l and not _is_internal(addr_l, internal_domains):
                            return addr_l
        except Exception:
            pass

    # fallback: none (internal sender)
    return None


@dataclass
class OutlookEnvironment:
    classic_available: bool
    new_outlook_detected: bool
    details: str


def detect_outlook_environment() -> OutlookEnvironment:
    new_outlook = False
    try:
        for candidate in ("olk.exe", "newoutlook.exe", "olks.exe"):
            proc = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"(Get-Process {candidate} -ErrorAction SilentlyContinue) -ne $null",
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode == 0 and proc.stdout.strip().lower() == "true":
                new_outlook = True
                break
    except Exception:
        pass

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        _ = outlook.GetNamespace("MAPI")
        version = safe_get(outlook, "Version", "unknown")
        return OutlookEnvironment(
            classic_available=True,
            new_outlook_detected=new_outlook,
            details=f"COM ok, version={version}",
        )
    except Exception as exc:
        detail = f"COM dispatch failed: {exc}"
        return OutlookEnvironment(
            classic_available=False,
            new_outlook_detected=new_outlook or True,
            details=detail,
        )
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class OutlookClient(OutlookLike):
    def __init__(self, cfg: AppConfig):
        self.cfg = cfg
        self.outlook = None
        self.ns = None

    def __enter__(self):
        env = detect_outlook_environment()
        if not env.classic_available:
            raise RuntimeError(
                f"Classic Outlook COM недоступен (details: {env.details}); New Outlook detected={env.new_outlook_detected}. Откройте Classic Outlook."
            )
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def get_folder(self):
        if not self.ns:
            raise RuntimeError("Outlook namespace not initialized")
        folder = None
        mailbox = self.cfg.qa_mailbox_override or self.cfg.mailbox
        if mailbox:
            folder = safe_call(lambda: self.ns.Folders[mailbox])
            if folder and self.cfg.folder:
                for part in self.cfg.folder.split("/"):
                    folder = safe_call(lambda f=folder, p=part: f.Folders[p])
                    if folder is None:
                        break
        if folder is None:
            folder = self.ns.GetDefaultFolder(FOLDER_INBOX)
        return folder

    def get_sent_folder(self):
        if not self.ns:
            raise RuntimeError("Outlook namespace not initialized")
        return self.ns.GetDefaultFolder(FOLDER_SENT)

    def iter_messages(
        self, since: datetime, until: Optional[datetime] = None
    ) -> Iterable:
        inbox = self.get_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        until_dt = until or datetime.now()
        restriction = f"[ReceivedTime] >= '{dt_to_restrict_str(since)}' AND [ReceivedTime] <= '{dt_to_restrict_str(until_dt)}'"
        restricted = items.Restrict(restriction)
        item = safe_call(lambda: restricted.GetFirst())
        while item:
            yield item
            item = safe_call(lambda: restricted.GetNext())

    def diagnose(self, days: int) -> Tuple[int, int, List[Tuple[str, int]]]:
        start = datetime.now() - timedelta(days=days)
        before_filter = 0
        after_filter = 0
        sender_counts: Counter[str] = Counter()
        for item in self.iter_messages(start):
            cls = safe_get(item, "Class", 0)
            if cls != 43:
                continue
            before_filter += 1
            sender = get_sender_smtp(item)
            sender_counts[sender] += 1
            if not passes_sender_filter(sender or "", self.cfg):
                continue
            after_filter += 1
        top10 = sender_counts.most_common(10)
        return before_filter, after_filter, top10

    def send_mail(
        self,
        subject: str,
        body: str,
        to: List[str],
        voting_options: Optional[str] = None,
        safe_only: bool = False,
        html_body: Optional[str] = None,
    ):
        mail = self.outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        if html_body:
            mail.HTMLBody = html_body
        mail.To = ";".join(to)
        if voting_options:
            mail.VotingOptions = voting_options
        if self.cfg.safe_mode or safe_only or not self.cfg.allow_send:
            mail.Display()
        else:
            try:
                mail.Send()
            except Exception as exc:
                log.warning("Send failed, fallback to Display: %s", exc)
                mail.Display()
        return mail

    def reply_overdue(
        self,
        original_entry_id: str,
        body: str,
        voting_options: Optional[str],
        html_body: Optional[str] = None,
    ) -> None:
        if not self.cfg.allow_send:
            safe_only = True
        else:
            safe_only = False
        try:
            item = self.ns.GetItemFromID(original_entry_id)
            reply = item.Reply()
            reply.Body = body + "\n\n" + reply.Body
            if html_body:
                reply.HTMLBody = html_body + reply.HTMLBody
            if voting_options:
                reply.VotingOptions = voting_options
            if self.cfg.safe_mode or safe_only:
                reply.Display()
            else:
                reply.Send()
        except Exception as exc:
            log.error("Failed to send reply for overdue: %s", exc)
            raise
