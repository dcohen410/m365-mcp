"""Settings for the server.

Settings come from environment variables first, then from a JSON file in the
user's config folder (written by `m365-mcp configure`). Nothing here is ever
committed to the repository: every organisation supplies its own client and
tenant IDs.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path

ENV_PREFIX = "M365_MCP_"

# Delegated Microsoft Graph permissions. The server only ever acts as the
# signed-in person, so it can reach exactly what they can reach in Outlook,
# OneDrive and SharePoint, and nothing more.
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
]

SEND_MODES = ("confirm", "disabled")


def config_dir() -> Path:
    override = os.environ.get(ENV_PREFIX + "CONFIG_DIR")
    if override:
        return Path(override).expanduser()
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        return base / "m365-mcp"
    return Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "m365-mcp"


def config_path() -> Path:
    return config_dir() / "config.json"


@dataclass
class Config:
    client_id: str = ""
    tenant_id: str = ""
    # Local folders the server may read from (for attachments/uploads) and
    # write to (for downloads). Anything outside these is refused.
    local_dirs: list[str] = field(default_factory=lambda: [str(Path.home() / "Downloads")])
    # Optional HTML file appended to the body of new drafts.
    signature_html: str = ""
    # "confirm": sending is allowed but every send is confirmed by the user.
    # "disabled": the server can only create drafts.
    send_mode: str = "confirm"

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    def is_configured(self) -> bool:
        return bool(self.client_id and self.tenant_id)

    def local_roots(self) -> list[Path]:
        return [Path(p).expanduser().resolve() for p in self.local_dirs if p]

    def signature(self) -> str:
        if not self.signature_html:
            return ""
        path = Path(self.signature_html).expanduser()
        return path.read_text(encoding="utf-8") if path.is_file() else ""


def load() -> Config:
    data: dict = {}
    path = config_path()
    if path.is_file():
        data = json.loads(path.read_text(encoding="utf-8"))

    cfg = Config(
        client_id=data.get("client_id", ""),
        tenant_id=data.get("tenant_id", ""),
        local_dirs=data.get("local_dirs") or Config().local_dirs,
        signature_html=data.get("signature_html", ""),
        send_mode=data.get("send_mode", "confirm"),
    )

    env = os.environ
    cfg.client_id = env.get(ENV_PREFIX + "CLIENT_ID", cfg.client_id)
    cfg.tenant_id = env.get(ENV_PREFIX + "TENANT_ID", cfg.tenant_id)
    if env.get(ENV_PREFIX + "LOCAL_DIRS"):
        cfg.local_dirs = [p for p in env[ENV_PREFIX + "LOCAL_DIRS"].split(os.pathsep) if p]
    cfg.signature_html = env.get(ENV_PREFIX + "SIGNATURE_HTML", cfg.signature_html)
    cfg.send_mode = env.get(ENV_PREFIX + "SEND_MODE", cfg.send_mode)
    if cfg.send_mode not in SEND_MODES:
        cfg.send_mode = "confirm"
    return cfg


def save(cfg: Config) -> Path:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "client_id": cfg.client_id,
                "tenant_id": cfg.tenant_id,
                "local_dirs": cfg.local_dirs,
                "signature_html": cfg.signature_html,
                "send_mode": cfg.send_mode,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return path
