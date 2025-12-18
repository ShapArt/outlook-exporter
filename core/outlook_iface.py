from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List, Protocol, Tuple


class OutlookLike(Protocol):
    """Protocol for Outlook client to allow fake injection in tests."""

    def iter_messages(self, since, until=None) -> Iterable: ...

    def send_mail(
        self,
        subject: str,
        body: str,
        to: List[str],
        voting_options: str | None = None,
        safe_only: bool = False,
        html_body: str | None = None,
    ): ...

    def reply_overdue(
        self,
        original_entry_id: str,
        body: str,
        voting_options: str | None,
        html_body: str | None = None,
    ) -> None: ...

    def diagnose(self, days: int) -> Tuple[int, int, List[Tuple[str, int]]]: ...


@dataclass
class FakeMail:
    Subject: str
    Body: str
    ReceivedTime: object
    Class: int = 43
    ConversationID: str | None = None
    EntryID: str | None = None
    VotingResponse: str | None = None
    ReplyTo: str | None = None
    SenderEmailAddress: str | None = None
    Sender: object | None = None
    ReplyRecipients: object | None = None


class FakeOutlookClient:
    """In-memory Outlook client for tests."""

    def __init__(self, messages: List[FakeMail] | None = None):
        self.messages = messages or []
        self.sent: List[dict] = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def iter_messages(self, since, until=None):
        return list(self.messages)

    def send_mail(
        self,
        subject: str,
        body: str,
        to: List[str],
        voting_options: str | None = None,
        safe_only: bool = False,
        html_body: str | None = None,
    ):
        self.sent.append(
            {
                "subject": subject,
                "body": body,
                "to": to,
                "voting": voting_options,
                "html": html_body,
            }
        )

    def reply_overdue(
        self,
        original_entry_id: str,
        body: str,
        voting_options: str | None,
        html_body: str | None = None,
    ) -> None:
        self.sent.append(
            {
                "reply_to": original_entry_id,
                "body": body,
                "voting": voting_options,
                "html": html_body,
            }
        )

    def diagnose(self, days: int):
        return len(self.messages), len(self.messages), []
