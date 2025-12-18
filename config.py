from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

APP_NAME = "NAOS_SLA_TRACKER"
# Store data in %APPDATA%/NAOS_SLA_TRACKER for portable exe and single data dir.
DEFAULT_APPDATA = Path(os.environ.get("APPDATA") or Path.cwd()) / APP_NAME


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _expand_path(value) -> Path:
    if isinstance(value, Path):
        return value
    return Path(os.path.expandvars(str(value))).expanduser()


@dataclass
class Paths:
    appdata_dir: Path = DEFAULT_APPDATA
    log_dir: Path = field(default_factory=lambda: DEFAULT_APPDATA / "logs")
    db_path: Path = field(default_factory=lambda: DEFAULT_APPDATA / "naos_sla.sqlite3")
    excel_path: Path = field(default_factory=lambda: DEFAULT_APPDATA / "tickets.xlsx")
    backup_dir: Path = field(default_factory=lambda: DEFAULT_APPDATA / "backups")
    orm_script_path: Path = field(
        default_factory=lambda: Path("NAOS_ORM_master script.xlsx")
    )

    def ensure(self) -> "Paths":
        self.appdata_dir = Path(self.appdata_dir)
        self.log_dir = Path(self.log_dir)
        self.db_path = Path(self.db_path)
        self.excel_path = Path(self.excel_path)
        self.backup_dir = Path(self.backup_dir)
        self.orm_script_path = Path(self.orm_script_path)
        _ensure_dir(self.appdata_dir)
        _ensure_dir(self.log_dir)
        _ensure_dir(self.backup_dir)
        return self


@dataclass
class AppConfig:
    version: str = "1.3.0"
    sender_filter: str = "support@ru.naos.com"
    sender_filter_mode: str = "off"  # off|contains|equals|domain
    sender_filter_value: Optional[str] = None
    mailbox: Optional[str] = None
    folder: Optional[str] = None
    lookback_days_open: int = 35
    ingest_days: int = 7
    overdue_days: int = 2
    timezone: str = "Europe/Moscow"
    safe_mode: bool = True
    allow_send: bool = False
    send_allow_domains: List[str] = field(
        default_factory=lambda: ["ru.naos.com", "naos.com"]
    )
    send_allowlist: List[str] = field(default_factory=list)
    test_allowlist: List[str] = field(
        default_factory=lambda: ["artem.shapovalov@ru.naos.com"]
    )
    wal_mode: bool = True
    excel_lock_timeout_sec: int = 5
    business_hours_start: int = 10
    business_hours_end: int = 19
    holidays: List[str] = field(default_factory=list)
    quiet_hours_start: int = 22
    quiet_hours_end: int = 8
    reminder_interval_hours: int = 24
    docs_url: Optional[str] = None
    sharepoint_url: Optional[str] = None
    excel_password: str = "naos"
    orm_similarity_days: int = 30
    orm_similarity_threshold: float = 0.42
    orm_max_suggestions: int = 3
    orm_script_path: Optional[str] = None
    qa_artem_email: Optional[str] = None
    qa_subject_prefix: str = "[QA][SLA]"
    qa_force_overdue_days: int = 5
    qa_mailbox_override: Optional[str] = None
    escalation_matrix: dict = field(
        default_factory=lambda: {
            "p1": [],
            "p2": [],
            "p3": [],
            "p4": [],
        }
    )
    status_catalog: List[str] = field(
        default_factory=lambda: [
            "new",
            "assigned",
            "responded",
            "resolved",
            "waiting_customer",
            "table",
            "otip",
            "overdue",
            "not_interesting",
        ]
    )
    status_hints: dict = field(
        default_factory=lambda: {
            "new": "Новое обращение, ещё не разобрали",
            "assigned": "Назначено/переслано",
            "responded": "Дан ответ",
            "resolved": "Закрыто/решено",
            "waiting_customer": "Ждём клиента",
            "table": "Требуются данные врача/таблица",
            "otip": "OTIP поток",
            "overdue": "Просрочка SLA",
            "not_interesting": "Неинтересно/спам/не наш",
        }
    )
    customer_internal_domains: List[str] = field(
        default_factory=lambda: ["ru.naos.com", "naos.com"]
    )
    sla_by_priority: dict = field(
        default_factory=lambda: {
            "p1": {"first_response_hours": 4, "resolution_hours": 24},
            "p2": {"first_response_hours": 8, "resolution_hours": 36},
            "p3": {"first_response_hours": 16, "resolution_hours": 48},
            "p4": {"first_response_hours": 24, "resolution_hours": 72},
        }
    )
    paths: Paths = field(default_factory=Paths)

    @classmethod
    def load(cls, path: Optional[Path] = None) -> "AppConfig":
        paths = Paths().ensure()
        cfg_path = path or (paths.appdata_dir / "config.json")
        if cfg_path.exists():
            try:
                data = json.loads(cfg_path.read_text(encoding="utf-8"))
                cfg = cls.from_dict(data)
                cfg.paths.ensure()
                cfg.excel_password = os.environ.get(
                    "NAOS_EXCEL_PASSWORD", cfg.excel_password
                )
                return cfg
            except Exception:
                pass
        cfg = cls(paths=paths.ensure())
        cfg.excel_password = os.environ.get("NAOS_EXCEL_PASSWORD", cfg.excel_password)
        cfg.save(cfg_path)
        return cfg

    @classmethod
    def from_dict(cls, data: dict) -> "AppConfig":
        paths_data = data.get("paths") or {}
        paths = Paths(
            appdata_dir=_expand_path(paths_data.get("appdata_dir", DEFAULT_APPDATA)),
            log_dir=_expand_path(paths_data.get("log_dir", DEFAULT_APPDATA / "logs")),
            db_path=_expand_path(
                paths_data.get("db_path", DEFAULT_APPDATA / "naos_sla.sqlite3")
            ),
            excel_path=_expand_path(
                paths_data.get("excel_path", DEFAULT_APPDATA / "tickets.xlsx")
            ),
            backup_dir=_expand_path(
                paths_data.get("backup_dir", DEFAULT_APPDATA / "backups")
            ),
            orm_script_path=_expand_path(
                paths_data.get("orm_script_path", "NAOS_ORM_master script.xlsx")
            ),
        ).ensure()
        return cls(
            version=data.get("version", "1.3.0"),
            sender_filter=data.get("sender_filter", "support@ru.naos.com"),
            sender_filter_mode=data.get("sender_filter_mode")
            or ("contains" if data.get("sender_filter") else "off"),
            sender_filter_value=data.get("sender_filter_value"),
            mailbox=data.get("mailbox"),
            folder=data.get("folder"),
            lookback_days_open=int(data.get("lookback_days_open", 35)),
            ingest_days=int(data.get("ingest_days", 7)),
            overdue_days=int(data.get("overdue_days", 2)),
            timezone=data.get("timezone", "Europe/Moscow"),
            safe_mode=bool(data.get("safe_mode", True)),
            allow_send=bool(data.get("allow_send", False)),
            send_allow_domains=list(
                data.get(
                    "send_allow_domains",
                    data.get("customer_internal_domains", ["ru.naos.com", "naos.com"]),
                )
            ),
            send_allowlist=list(data.get("send_allowlist", [])),
            test_allowlist=list(
                data.get("test_allowlist", ["artem.shapovalov@ru.naos.com"])
            ),
            wal_mode=bool(data.get("wal_mode", True)),
            excel_lock_timeout_sec=int(data.get("excel_lock_timeout_sec", 5)),
            business_hours_start=int(data.get("business_hours_start", 10)),
            business_hours_end=int(data.get("business_hours_end", 19)),
            holidays=list(data.get("holidays", [])),
            quiet_hours_start=int(data.get("quiet_hours_start", 22)),
            quiet_hours_end=int(data.get("quiet_hours_end", 8)),
            reminder_interval_hours=int(data.get("reminder_interval_hours", 24)),
            docs_url=data.get("docs_url"),
            sharepoint_url=data.get("sharepoint_url"),
            excel_password=str(data.get("excel_password", "naos")),
            orm_similarity_days=int(data.get("orm_similarity_days", 30)),
            orm_similarity_threshold=float(data.get("orm_similarity_threshold", 0.42)),
            orm_max_suggestions=int(data.get("orm_max_suggestions", 3)),
            orm_script_path=data.get("orm_script_path"),
            qa_artem_email=data.get("qa_artem_email"),
            qa_subject_prefix=data.get("qa_subject_prefix", "[QA][SLA]"),
            qa_force_overdue_days=int(data.get("qa_force_overdue_days", 5)),
            qa_mailbox_override=data.get("qa_mailbox_override"),
            escalation_matrix=dict(data.get("escalation_matrix", {}))
            or {"p1": [], "p2": [], "p3": [], "p4": []},
            sla_by_priority=dict(data.get("sla_by_priority", {}))
            or {
                "p1": {"first_response_hours": 4, "resolution_hours": 24},
                "p2": {"first_response_hours": 8, "resolution_hours": 36},
                "p3": {"first_response_hours": 16, "resolution_hours": 48},
                "p4": {"first_response_hours": 24, "resolution_hours": 72},
            },
            status_catalog=list(data.get("status_catalog", []))
            or [
                "new",
                "assigned",
                "responded",
                "resolved",
                "waiting_customer",
                "table",
                "otip",
                "overdue",
                "not_interesting",
            ],
            status_hints=dict(data.get("status_hints", {}))
            or {
                "new": "Новое обращение, ещё не разобрали",
                "assigned": "Назначено/переслано",
                "responded": "Дан ответ",
                "resolved": "Закрыто/решено",
                "waiting_customer": "Ждём клиента",
                "table": "Требуются данные врача/таблица",
                "otip": "OTIP поток",
                "overdue": "Просрочка SLA",
                "not_interesting": "Неинтересно/спам/не наш",
            },
            customer_internal_domains=list(
                data.get("customer_internal_domains", ["ru.naos.com", "naos.com"])
            ),
            paths=paths,
        )

    def save(self, path: Optional[Path] = None) -> None:
        cfg_path = path or (self.paths.appdata_dir / "config.json")
        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        cfg_path.write_text(
            json.dumps(
                {
                    "version": self.version,
                    "sender_filter": self.sender_filter,
                    "sender_filter_mode": self.sender_filter_mode,
                    "sender_filter_value": self.sender_filter_value,
                    "mailbox": self.mailbox,
                    "folder": self.folder,
                    "lookback_days_open": self.lookback_days_open,
                    "ingest_days": self.ingest_days,
                    "overdue_days": self.overdue_days,
                    "timezone": self.timezone,
                    "safe_mode": self.safe_mode,
                    "allow_send": self.allow_send,
                    "send_allow_domains": self.send_allow_domains,
                    "send_allowlist": self.send_allowlist,
                    "test_allowlist": self.test_allowlist,
                    "wal_mode": self.wal_mode,
                    "excel_lock_timeout_sec": self.excel_lock_timeout_sec,
                    "business_hours_start": self.business_hours_start,
                    "business_hours_end": self.business_hours_end,
                    "holidays": self.holidays,
                    "quiet_hours_start": self.quiet_hours_start,
                    "quiet_hours_end": self.quiet_hours_end,
                    "reminder_interval_hours": self.reminder_interval_hours,
                    "docs_url": self.docs_url,
                    "sharepoint_url": self.sharepoint_url,
                    "excel_password": self.excel_password,
                    "orm_similarity_days": self.orm_similarity_days,
                    "orm_similarity_threshold": self.orm_similarity_threshold,
                    "orm_max_suggestions": self.orm_max_suggestions,
                    "orm_script_path": self.orm_script_path
                    or str(self.paths.orm_script_path),
                    "qa_artem_email": self.qa_artem_email,
                    "qa_subject_prefix": self.qa_subject_prefix,
                    "qa_force_overdue_days": self.qa_force_overdue_days,
                    "qa_mailbox_override": self.qa_mailbox_override,
                    "escalation_matrix": self.escalation_matrix,
                    "sla_by_priority": self.sla_by_priority,
                    "status_catalog": self.status_catalog,
                    "status_hints": self.status_hints,
                    "customer_internal_domains": self.customer_internal_domains,
                    "paths": {
                        "appdata_dir": str(self.paths.appdata_dir),
                        "log_dir": str(self.paths.log_dir),
                        "db_path": str(self.paths.db_path),
                        "excel_path": str(self.paths.excel_path),
                        "backup_dir": str(self.paths.backup_dir),
                        "orm_script_path": str(self.paths.orm_script_path),
                    },
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )


def default_config() -> AppConfig:
    return AppConfig.load()
