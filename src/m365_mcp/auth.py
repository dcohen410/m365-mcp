"""Microsoft sign-in with MSAL.

The app registration is a public client (no secret). Tokens are cached in
the operating system's secure store: Keychain on macOS, DPAPI on Windows and
libsecret on Linux. On a Linux machine without libsecret the cache falls back
to a file only readable by the current user, and a warning is printed.
"""

from __future__ import annotations

import logging
import os
import sys

import msal
from msal_extensions import FilePersistence, PersistedTokenCache, build_encrypted_persistence

from . import config as config_mod
from .config import SCOPES, Config

log = logging.getLogger(__name__)


class NotSignedIn(Exception):
    pass


def _token_cache() -> PersistedTokenCache:
    path = config_mod.config_dir() / "token_cache.bin"
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        persistence = build_encrypted_persistence(str(path))
    except Exception as exc:  # pragma: no cover - depends on the OS keyring
        if sys.platform in ("darwin", "win32"):
            raise
        print(
            f"Warning: no secure keyring available ({exc}); "
            f"storing the sign-in token in {path} readable only by you.",
            file=sys.stderr,
        )
        persistence = FilePersistence(str(path))
        if path.exists():
            os.chmod(path, 0o600)
    return PersistedTokenCache(persistence)


def _app(cfg: Config) -> msal.PublicClientApplication:
    if not cfg.is_configured():
        raise NotSignedIn(
            "m365-mcp is not configured. Run `m365-mcp configure --client-id <id> --tenant-id <id>` "
            "(see the README for how to get these)."
        )
    return msal.PublicClientApplication(
        cfg.client_id, authority=cfg.authority, token_cache=_token_cache()
    )


def login(cfg: Config, device_code: bool = False) -> str:
    """Interactive sign-in. Returns the signed-in username."""
    app = _app(cfg)
    if device_code:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not start device sign-in: {flow.get('error_description', flow)}")
        print(flow["message"], file=sys.stderr)
        result = app.acquire_token_by_device_flow(flow)
    else:
        result = app.acquire_token_interactive(scopes=SCOPES, prompt="select_account")
    if "access_token" not in result:
        raise RuntimeError(f"Sign-in failed: {result.get('error_description') or result.get('error')}")
    return result.get("id_token_claims", {}).get("preferred_username", "")


def logout(cfg: Config) -> None:
    app = _app(cfg)
    for account in app.get_accounts():
        app.remove_account(account)


def signed_in_user(cfg: Config) -> str | None:
    accounts = _app(cfg).get_accounts()
    return accounts[0].get("username") if accounts else None


class TokenProvider:
    """Hands out access tokens silently, refreshing them as needed."""

    def __init__(self, cfg: Config):
        self._cfg = cfg
        self._app: msal.PublicClientApplication | None = None

    def __call__(self) -> str:
        if self._app is None:
            self._app = _app(self._cfg)
        accounts = self._app.get_accounts()
        if not accounts:
            raise NotSignedIn("Not signed in to Microsoft 365. Run `m365-mcp login` in a terminal, then try again.")
        result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result or "access_token" not in result:
            raise NotSignedIn(
                "Your Microsoft 365 sign-in has expired or needs new permissions. "
                "Run `m365-mcp login` in a terminal, then try again."
            )
        return result["access_token"]
