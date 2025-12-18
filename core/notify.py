from __future__ import annotations

from datetime import datetime
from typing import List, Optional, Tuple

from config import AppConfig
from core.logger import get_logger
from core.outlook import OutlookClient

log = get_logger(__name__)

VOTING_OPTIONS_RU = "OK;Нужно время;Закрыть;Не наш"


def _html_section(title: str, content: str) -> str:
    return f"""
    <div style="margin-bottom:12px;">
      <div style="font-weight:600;color:#0F172A;margin-bottom:4px;">{title}</div>
      <div style="color:#111827;line-height:1.4;">{content}</div>
    </div>
    """


def _build_overdue_html(
    cfg: AppConfig,
    ticket_subject: str,
    ticket_id: int,
    priority: Optional[str],
    excel_path: str,
) -> str:
    prio_tag = priority.upper() if priority else "n/a"
    guide_url = cfg.docs_url or cfg.sharepoint_url or ""
    guide_link = (
        f'<a href="{guide_url}">Гайд/шпаргалка</a>'
        if guide_url
        else "Гайд в SharePoint"
    )
    excel_hint = f"Excel: {excel_path}"
    intro = f"<div style='font-size:18px;font-weight:700;color:#111827;'>Просрочка SLA #{ticket_id} — {prio_tag}</div>"
    what_to_do = """
    1) Нажмите Voting: <b>OK / Нужно время / Закрыть / Не наш</b><br>
    2) Или ответьте командами: <code>/status</code>, <code>/prio</code>, <code>/owner</code>, <code>/comment</code>
    """
    links = f"{guide_link}<br>{excel_hint}"
    return f"""
    <html>
      <body style="font-family:Segoe UI, Arial, sans-serif; background:#F8FAFC; padding:16px;">
        <div style="max-width:720px; margin:0 auto; background:#fff; padding:16px; border-radius:10px; box-shadow:0 4px 14px rgba(0,0,0,0.06);">
          {intro}
          <div style="color:#334155;margin:8px 0 12px 0;">{ticket_subject}</div>
          {_html_section("Что сделать сейчас", what_to_do)}
          {_html_section("Ссылки и гайд", links)}
          {_html_section("Данные обращения", f"ID: <b>{ticket_id}</b><br>Тема: {ticket_subject}<br>Приоритет: {prio_tag}<br>Сформировано: {datetime.utcnow().isoformat()}Z")}
        </div>
      </body>
    </html>
    """


def _filter_recipients(
    cfg: AppConfig, recipients: List[str]
) -> Tuple[List[str], List[str]]:
    allowed_domains = {d.lower().strip() for d in (cfg.send_allow_domains or []) if d}
    allowlist = {a.lower().strip() for a in (cfg.send_allowlist or []) if a}
    allowlist |= {a.lower().strip() for a in (cfg.test_allowlist or []) if a}
    allowed: List[str] = []
    blocked: List[str] = []
    for rec in recipients:
        r = (rec or "").strip().lower()
        if not r:
            continue
        if r in allowlist:
            allowed.append(r)
            continue
        if allowed_domains and any(
            r.endswith("@" + d) or r.endswith("." + d) for d in allowed_domains
        ):
            allowed.append(r)
            continue
        blocked.append(r)
    return allowed, blocked


def send_test_mail(cfg: AppConfig, subject: str, body: str) -> None:
    to_list = cfg.test_allowlist or []
    if not to_list:
        raise ValueError(
            "Test allowlist is empty; configure test_allowlist in config.json"
        )
    allowed, blocked = _filter_recipients(cfg, to_list)
    if blocked:
        log.info("Blocked test recipients by policy: %s", blocked)
    if not allowed:
        raise ValueError("No allowed test recipients after policy filtering")
    with OutlookClient(cfg) as outlook:
        outlook.send_mail(
            subject,
            body,
            allowed,
            voting_options=VOTING_OPTIONS_RU,
            safe_only=cfg.safe_mode,
            html_body=f"<p>{body}</p>",
        )
    log.info("Test mail queued to %s", allowed)


def send_overdue_mail(
    cfg: AppConfig,
    responsible: str,
    ticket_subject: str,
    ticket_id: int,
    priority: str | None = None,
    original_entry_id: str | None = None,
    preview_only: bool = False,
):
    prefix = "[SLA]"
    prio = priority or ""
    prio_tag = f"[{prio.upper()}]" if prio else ""
    subj = f"{prefix}{prio_tag} Просрочка #{ticket_id}: {ticket_subject}"
    plain = (
        "Просрочка SLA.\n\n"
        f"ID: {ticket_id}\n"
        f"Тема: {ticket_subject}\n"
        f"Приоритет: {prio or 'n/a'}\n"
        f"Время: {datetime.utcnow().isoformat()}Z\n"
        "Ответьте голосованием (OK / Нужно время / Закрыть / Не наш)\n"
        "или командами /status /prio /owner /comment.\n"
    )
    html = _build_overdue_html(
        cfg, ticket_subject, ticket_id, priority, str(cfg.paths.excel_path)
    )
    voting_opts = VOTING_OPTIONS_RU
    to_list = [responsible]
    if priority and cfg.escalation_matrix.get(priority):
        to_list += cfg.escalation_matrix.get(priority, [])
    allowed, blocked = _filter_recipients(cfg, to_list)
    quiet = False
    try:
        from core.sla import is_quiet_hours

        quiet = is_quiet_hours(cfg)
    except Exception:
        quiet = False
    if blocked:
        log.info("Blocked recipients by policy: %s", blocked)
    if preview_only or cfg.safe_mode or not cfg.allow_send or quiet:
        log.info("PREVIEW send to %s: %s (quiet=%s)", allowed, subj, quiet)
        return
    if not allowed:
        log.info(
            "No allowed recipients for ticket %s after policy filtering", ticket_id
        )
        return
    with OutlookClient(cfg) as outlook:
        if original_entry_id:
            try:
                outlook.reply_overdue(
                    original_entry_id, plain, voting_options=voting_opts, html_body=html
                )
                log.info(
                    "Sent reply-based overdue notification for ticket %s", ticket_id
                )
                return
            except Exception as exc:
                log.warning("Reply-based send failed, fallback to new mail: %s", exc)
        outlook.send_mail(
            subj, plain, allowed, voting_options=voting_opts, html_body=html
        )
        log.info("Sent new-mail overdue notification for ticket %s", ticket_id)
