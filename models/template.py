from dataclasses import dataclass, field
from typing import List, Optional, Dict
from datetime import datetime


@dataclass
class Template:
    """
    Email template model.

    subject: subject template with placeholders like {{stopover_code}}
    body: body template with placeholders
    placeholders: list of supported placeholders for preview/help
    """
    subject: str = "Stopover Report - {{stopover_code}}"
    body: str = (
        "Dear Team,\n\n"
        "Please find attached the stopover report for {{stopover_code}}.\n\n"
        "Best regards,\n"
        "PDF Stopover Analyzer"
    )
    placeholders: List[str] = field(default_factory=lambda: ["{{stopover_code}}"])


@dataclass
class Account:
    """
    Connected account information for display and sending context.
    """
    id: Optional[str] = None
    name: Optional[str] = None
    email: Optional[str] = None
    provider: str = "Outlook"
    connected: bool = False


class SendStatus:
    """
    Simple status constants for send operations.
    """
    PENDING = "PENDING"
    SENDING = "SENDING"
    SUCCESS = "SUCCESS"
    FAILED = "FAILED"
    QUEUED = "QUEUED"


@dataclass
class StopoverMeta:
    """
    Metadata shown on each stopover item.
    """
    stopover_code: str
    status: str = SendStatus.PENDING
    last_sent_time: Optional[datetime] = None
    issues: List[str] = field(default_factory=list)

    def to_display_dict(self) -> Dict[str, str]:
        return {
            "code": self.stopover_code,
            "status": self.status,
            "last_sent": self.last_sent_time.isoformat() if self.last_sent_time else "Never",
            "issues": "; ".join(self.issues) if self.issues else "",
        }