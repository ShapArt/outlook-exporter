import types

from core.outlook import extract_customer_email


class FakeRecip:
    def __init__(self, address):
        self.Address = address


def test_external_sender_taken():
    email = extract_customer_email(
        msg=None,
        sender_smtp="user@example.com",
        body_text="",
        subject="",
        internal_domains=["ru.naos.com"],
    )
    assert email == "user@example.com"


def test_forward_block_in_body():
    body = "From: client@domain.com\nSubject: Hi\nBody text"
    email = extract_customer_email(
        msg=None,
        sender_smtp="agent@ru.naos.com",
        body_text=body,
        subject="",
        internal_domains=["ru.naos.com"],
    )
    assert email == "client@domain.com"


def test_reply_recipients_used():
    msg = types.SimpleNamespace()
    msg.ReplyRecipients = [FakeRecip("customer@ext.com")]
    email = extract_customer_email(
        msg,
        sender_smtp="agent@ru.naos.com",
        body_text="",
        subject="",
        internal_domains=["ru.naos.com"],
    )
    assert email == "customer@ext.com"


def test_reply_to_field_used():
    msg = types.SimpleNamespace()
    msg.ReplyTo = "customer2@ext.com"
    email = extract_customer_email(
        msg,
        sender_smtp="agent@ru.naos.com",
        body_text="",
        subject="",
        internal_domains=["ru.naos.com"],
    )
    assert email == "customer2@ext.com"


def test_internal_only_returns_none():
    email = extract_customer_email(
        msg=None,
        sender_smtp="agent@ru.naos.com",
        body_text="No emails here",
        subject="",
        internal_domains=["ru.naos.com"],
    )
    assert email is None
