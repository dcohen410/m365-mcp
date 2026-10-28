"""Command line: `m365-mcp configure | login | logout | status | serve`."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import __version__, auth, mail
from . import config as config_mod

SIGNATURE_IMAGE_NAME = "signature-image"
SIGNATURE_IMAGE_LIMIT = 1024 * 1024


def _configure(args: argparse.Namespace) -> int:
    cfg = config_mod.load()
    if args.client_id:
        cfg.client_id = args.client_id.strip()
    if args.tenant_id:
        cfg.tenant_id = args.tenant_id.strip()
    if args.local_dir:
        cfg.local_dirs = [os.path.expanduser(p) for p in args.local_dir]
    if args.signature is not None:
        cfg.signature_html = os.path.expanduser(args.signature)
    if args.signature_image is not None:
        try:
            cfg.signature_image = _store_signature_image(args.signature_image)
        except ValueError as exc:
            print(exc, file=sys.stderr)
            return 1
    if args.send_mode:
        cfg.send_mode = args.send_mode
    if args.read_only is not None:
        cfg.read_only = args.read_only
    if args.tools:
        try:
            cfg.tool_groups = config_mod.parse_tool_groups(args.tools)
        except ValueError as exc:
            print(exc, file=sys.stderr)
            return 1
    if args.delete_mode:
        cfg.delete_mode = args.delete_mode
    path = config_mod.save(cfg)
    print(f"Saved settings to {path}")
    if not cfg.is_configured():
        print("Still needed: --client-id and --tenant-id (from your Entra app registration).")
        return 1
    print("Next: run `m365-mcp login`.")
    return 0


def _store_signature_image(source: str) -> str:
    """Copy the image into the settings folder, so moving or deleting the original
    doesn't lose it. Returns the stored path, or "" to remove the image."""
    path = Path(source).expanduser() if source else None
    if path:
        if not path.is_file():
            raise ValueError(f"Can't find the signature image {path}.")
        if path.suffix.lower() not in mail.SIGNATURE_IMAGE_TYPES:
            raise ValueError("The signature image must be a PNG, JPEG or GIF.")
        if path.stat().st_size > SIGNATURE_IMAGE_LIMIT:
            raise ValueError("The signature image is over 1 MB. Please use a smaller one.")
        data = path.read_bytes()
    for old in config_mod.config_dir().glob(SIGNATURE_IMAGE_NAME + ".*"):
        old.unlink()
    if not path:
        return ""
    stored = config_mod.config_dir() / (SIGNATURE_IMAGE_NAME + path.suffix.lower())
    stored.parent.mkdir(parents=True, exist_ok=True)
    stored.write_bytes(data)
    return str(stored)


def _login(args: argparse.Namespace) -> int:
    user = auth.login(config_mod.load(), device_code=args.device_code)
    print(f"Signed in as {user}. You can close the browser tab.")
    return 0


def _logout(_: argparse.Namespace) -> int:
    auth.logout(config_mod.load())
    print("Signed out and removed the saved sign-in.")
    return 0


def _status(_: argparse.Namespace) -> int:
    cfg = config_mod.load()
    print(f"m365-mcp {__version__}")
    print(f"Settings file:  {config_mod.config_path()}")
    print(f"Client ID:      {cfg.client_id or '(not set)'}")
    print(f"Tenant ID:      {cfg.tenant_id or '(not set)'}")
    print(f"Local folders:  {', '.join(cfg.local_dirs) or '(none)'}")
    print(f"Signature:      {cfg.signature_html or '(none)'}")
    print(f"Signature image: {cfg.signature_image or '(none)'}")
    print(f"Sending:        {cfg.send_mode}")
    print(f"Read-only:      {'yes' if cfg.read_only else 'no'}")
    print(f"Tool groups:    {', '.join(cfg.tool_groups)}")
    print(f"Deleting files: {cfg.delete_mode}")
    if cfg.is_configured():
        print(f"Signed in as:   {auth.signed_in_user(cfg) or '(not signed in; run m365-mcp login)'}")
    return 0


def _serve(_: argparse.Namespace) -> int:
    from .server import run

    run()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="m365-mcp", description=__doc__)
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command")

    p = sub.add_parser("configure", help="save your organisation's app registration and options")
    p.add_argument("--client-id", help="Application (client) ID of your Entra app registration")
    p.add_argument("--tenant-id", help="Directory (tenant) ID")
    p.add_argument("--local-dir", action="append",
                   help="a folder on this computer the server may read and write (repeatable; default ~/Downloads)")
    p.add_argument("--signature", help="path to an HTML email signature to add to new drafts ('' to remove)")
    p.add_argument("--signature-image",
                   help="a PNG, JPEG or GIF signature to embed at the end of new drafts; it's copied into the "
                        "settings folder ('' to remove)")
    p.add_argument("--send-mode", choices=config_mod.SEND_MODES,
                   help="'confirm' (default): sending allowed with confirmation; 'disabled': drafts only")
    p.add_argument("--read-only", action=argparse.BooleanOptionalAction, default=None,
                   help="only offer tools that read; nothing can be moved, written, drafted or sent")
    p.add_argument("--tools", help=f"comma-separated tool groups to offer: {', '.join(config_mod.TOOL_GROUPS)} "
                                   "or 'all' (default all)")
    p.add_argument("--delete-mode", choices=config_mod.DELETE_MODES,
                   help="'confirm' (default): deleting files to the recycle bin allowed with confirmation; "
                        "'disabled': no deleting")
    p.set_defaults(func=_configure)

    p = sub.add_parser("login", help="sign in to Microsoft 365 in your browser")
    p.add_argument("--device-code", action="store_true", help="sign in with a code instead of opening a browser")
    p.set_defaults(func=_login)

    sub.add_parser("logout", help="sign out and delete the saved sign-in").set_defaults(func=_logout)
    sub.add_parser("status", help="show settings and who is signed in").set_defaults(func=_status)
    sub.add_parser("serve", help="run the MCP server over stdio (what Claude runs)").set_defaults(func=_serve)

    args = parser.parse_args(argv)
    if not getattr(args, "func", None):
        args.func = _serve
    try:
        return args.func(args)
    except (auth.NotSignedIn, RuntimeError) as exc:
        print(exc, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
