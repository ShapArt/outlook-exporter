from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Callable, Optional

DEFAULT_MAX_BYTES = 2 * 1024 * 1024
DEFAULT_BACKUP_COUNT = 3
_ui_sinks: list[Callable[[str], None]] = []


class CallbackHandler(logging.Handler):
    """Lightweight bridge for piping logs into UI callbacks."""

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
        except Exception:
            msg = record.getMessage()
        for cb in list(_ui_sinks):
            try:
                cb(msg)
            except Exception:
                # Do not propagate UI errors back to logging stack.
                pass


def setup_logging(log_dir: Path, level: str = "INFO") -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "outlook_sla.log"

    logger = logging.getLogger("naos_sla")
    if logger.handlers:
        return logger

    logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = RotatingFileHandler(
        log_path,
        maxBytes=DEFAULT_MAX_BYTES,
        backupCount=DEFAULT_BACKUP_COUNT,
        encoding="utf-8",
    )
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    cb = CallbackHandler()
    cb.setFormatter(formatter)
    logger.addHandler(cb)

    logger.debug("Logger initialized at %s", log_path)
    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    base = logging.getLogger("naos_sla")
    return base if name is None else base.getChild(name)


def register_ui_sink(callback: Callable[[str], None]) -> None:
    if callback not in _ui_sinks:
        _ui_sinks.append(callback)


def clear_ui_sinks() -> None:
    _ui_sinks.clear()
