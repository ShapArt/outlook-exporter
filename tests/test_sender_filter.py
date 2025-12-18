from config import AppConfig
from core.utils import passes_sender_filter


def _cfg(mode="off", value=None, legacy=None):
    cfg = AppConfig()
    cfg.sender_filter_mode = mode
    cfg.sender_filter_value = value
    if legacy is not None:
        cfg.sender_filter = legacy
    return cfg


def test_filter_off_allows_any():
    assert passes_sender_filter("a@b.c", _cfg("off", "x")) is True


def test_filter_contains():
    cfg = _cfg("contains", "naos.com")
    assert passes_sender_filter("user@ru.naos.com", cfg) is True
    assert passes_sender_filter("user@other.com", cfg) is False


def test_filter_equals():
    cfg = _cfg("equals", "person@example.com")
    assert passes_sender_filter("person@example.com", cfg) is True
    assert passes_sender_filter("other@example.com", cfg) is False


def test_filter_domain():
    cfg = _cfg("domain", "example.com")
    assert passes_sender_filter("user@example.com", cfg) is True
    assert passes_sender_filter("user@mail.example.com", cfg) is True
    assert passes_sender_filter("user@other.com", cfg) is False
