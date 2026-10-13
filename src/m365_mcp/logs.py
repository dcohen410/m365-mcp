"""Keep sign-in tokens and email addresses out of the server's logs.

MCP clients usually save a server's stderr to a log file, so anything logged
there should be safe to paste into a bug report.
"""

from __future__ import annotations

import logging
import re

_PATTERNS = [
    # Authorization headers.
    (re.compile(r"(?i)\bBearer\s+[A-Za-z0-9._~+/=-]+"), "Bearer [redacted]"),
    # JSON web tokens (access and ID tokens).
    (re.compile(r"\beyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]*"), "[redacted token]"),
    # Named secrets in JSON, form bodies or query strings.
    (re.compile(r"(?i)\b(access_token|refresh_token|id_token|client_secret|client_info)"
                r"([\"']?\s*[:=]\s*[\"']?)[^\s\"'&,}]+"), r"\1\2[redacted]"),
    # Email addresses.
    (re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+"), "[email]"),
]


def redact(text: str) -> str:
    for pattern, replacement in _PATTERNS:
        text = pattern.sub(replacement, text)
    return text


class RedactingFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        try:
            message = record.getMessage()
        except Exception:
            message = str(record.msg)
        record.msg, record.args = redact(message), None
        if record.exc_info and not record.exc_text:
            record.exc_text = logging.Formatter().formatException(record.exc_info)
        if record.exc_text:
            record.exc_text = redact(record.exc_text)
            record.exc_info = None
        return True


_FILTER = RedactingFilter()


def install() -> None:
    """Redact everything the root logger's handlers write, and stop the HTTP
    libraries logging full request URLs (which carry search text)."""
    root = logging.getLogger()
    if not root.handlers:
        logging.basicConfig(level=logging.WARNING)
    for handler in root.handlers:
        if _FILTER not in handler.filters:
            handler.addFilter(_FILTER)
    for name in ("httpx", "httpcore", "msal"):
        logging.getLogger(name).setLevel(logging.WARNING)
