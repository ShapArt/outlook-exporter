from config import AppConfig
from core.utils import passes_sender_filter


def test_sender_filter_off():
    cfg = AppConfig()
    cfg.sender_filter_mode = "off"
    assert passes_sender_filter("any@example.com", cfg) is True


def test_sender_filter_contains():
    cfg = AppConfig()
    cfg.sender_filter_mode = "contains"
    cfg.sender_filter_value = "naos.com"
    assert passes_sender_filter("user@naos.com", cfg) is True
    assert passes_sender_filter("user@other.com", cfg) is False


def test_sender_filter_equals_case_insensitive():
    cfg = AppConfig()
    cfg.sender_filter_mode = "equals"
    cfg.sender_filter_value = "user@naos.com"
    assert passes_sender_filter("User@Naos.com", cfg) is True
    assert passes_sender_filter("other@naos.com", cfg) is False


def test_sender_filter_domain():
    cfg = AppConfig()
    cfg.sender_filter_mode = "domain"
    cfg.sender_filter_value = "naos.com"
    assert passes_sender_filter("user@sub.naos.com", cfg) is True
    assert passes_sender_filter("user@example.com", cfg) is False
