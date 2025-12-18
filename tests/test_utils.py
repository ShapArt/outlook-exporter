from core.utils import clean_text, normalize_subject


def test_normalize_subject_removes_prefixes():
    assert normalize_subject("Re: Fw: FW: hello") == "hello"
    assert normalize_subject("") == ""


def test_clean_text_strips_quote_headers():
    src = "From: someone\nSubject: hi\n> quoted\nBody line\n\n"
    assert clean_text(src) == "Body line"
