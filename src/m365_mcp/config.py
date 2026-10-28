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
# OneDrive and SharePoint, and nothing more. MailboxSettings.Read is only for
# the mailbox's time zone, so calendar times match what Outlook shows.
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "Calendars.ReadWrite",
    "MailboxSettings.Read",
]

SEND_MODES = ("confirm", "disabled")
DELETE_MODES = ("confirm", "disabled")

# Tools come in groups that can be switched off independently.
TOOL_GROUPS = ("mail", "drafts", "files", "excel", "calendar")
# Groups that existed before "calendar". A settings file that lists all of these
# was written when they meant "everything", so newer groups are switched on too.
_ORIGINAL_GROUPS = ("mail", "drafts", "files", "excel")


def _truthy(value: str) -> bool:
    return value.strip().lower() in ("1", "true", "yes", "on")


def parse_tool_groups(value: str | list[str]) -> list[str]:
    """Turn "mail,files" (or "all") into a list of known group names."""
    items = value.split(",") if isinstance(value, str) else value
    names = [i.strip().lower() for i in items if i and i.strip()]
    if not names or "all" in names:
        return list(TOOL_GROUPS)
    unknown = [n for n in names if n not in TOOL_GROUPS]
    if unknown:
        raise ValueError(f"Unknown tool group(s): {', '.join(unknown)}. Choose from: {', '.join(TOOL_GROUPS)}.")
    return [g for g in TOOL_GROUPS if g in names]


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
    # Optional image (PNG, JPEG or GIF) embedded at the end of new drafts,
    # after signature_html. `configure` copies it into the settings folder.
    signature_image: str = ""
    # "confirm": sending is allowed but every send is confirmed by the user.
    # "disabled": the server can only create drafts.
    send_mode: str = "confirm"
    # Only offer tools that read, never ones that change mail or files.
    read_only: bool = False
    # Which groups of tools to offer (see TOOL_GROUPS).
    tool_groups: list[str] = field(default_factory=lambda: list(TOOL_GROUPS))
    # "confirm": files can be deleted (to the recycle bin), confirmed by the user each time.
    # "disabled": no deleting.
    delete_mode: str = "confirm"

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

    def signature_image_file(self) -> Path | None:
        if not self.signature_image:
            return None
        path = Path(self.signature_image).expanduser()
        return path if path.is_file() else None


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
        signature_image=data.get("signature_image", ""),
        send_mode=data.get("send_mode", "confirm"),
        read_only=bool(data.get("read_only", False)),
        tool_groups=_saved_groups(data.get("tool_groups")),
        delete_mode=data.get("delete_mode", "confirm"),
    )

    env = os.environ
    cfg.client_id = env.get(ENV_PREFIX + "CLIENT_ID", cfg.client_id)
    cfg.tenant_id = env.get(ENV_PREFIX + "TENANT_ID", cfg.tenant_id)
    if env.get(ENV_PREFIX + "LOCAL_DIRS"):
        cfg.local_dirs = [p for p in env[ENV_PREFIX + "LOCAL_DIRS"].split(os.pathsep) if p]
    cfg.signature_html = env.get(ENV_PREFIX + "SIGNATURE_HTML", cfg.signature_html)
    cfg.signature_image = env.get(ENV_PREFIX + "SIGNATURE_IMAGE", cfg.signature_image)
    cfg.send_mode = env.get(ENV_PREFIX + "SEND_MODE", cfg.send_mode)
    if cfg.send_mode not in SEND_MODES:
        cfg.send_mode = "confirm"
    if env.get(ENV_PREFIX + "READ_ONLY"):
        cfg.read_only = _truthy(env[ENV_PREFIX + "READ_ONLY"])
    if env.get(ENV_PREFIX + "TOOLS"):
        cfg.tool_groups = _groups_or_default(env[ENV_PREFIX + "TOOLS"])
    cfg.delete_mode = env.get(ENV_PREFIX + "DELETE_MODE", cfg.delete_mode)
    if cfg.delete_mode not in DELETE_MODES:
        cfg.delete_mode = "disabled"
    return cfg


def _groups_or_default(value) -> list[str]:
    # A bad value in the settings should not stop the server starting, so fall
    # back to every group; `configure` rejects bad values before saving them.
    try:
        return parse_tool_groups(value or "all")
    except ValueError:
        return list(TOOL_GROUPS)


def _saved_groups(value) -> list[str]:
    groups = _groups_or_default(value)
    if set(_ORIGINAL_GROUPS) <= set(groups):
        return list(TOOL_GROUPS)
    return groups


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
                "signature_image": cfg.signature_image,
                "send_mode": cfg.send_mode,
                "read_only": cfg.read_only,
                "tool_groups": cfg.tool_groups,
                "delete_mode": cfg.delete_mode,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return path
