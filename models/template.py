from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional




class SendStatus:
    """Simple status constants for send operations."""
    PENDING = "PENDING"
    SENDING = "SENDING"
    SUCCESS = "SUCCESS"
    FAILED = "FAILED"
    QUEUED = "QUEUED"


@dataclass
class StopoverMeta:
    """Metadata shown on each stopover item."""
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