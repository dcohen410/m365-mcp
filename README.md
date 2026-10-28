# m365-mcp

A local [MCP](https://modelcontextprotocol.io) server that lets Claude work with your **Outlook mail and
calendar, OneDrive and SharePoint** through Microsoft Graph. It runs on your own computer and signs in as you,
so it can only reach what you can already reach.

It's built to fill the gaps in typical Microsoft 365 connectors:

- **Save email attachments** straight into a OneDrive or SharePoint folder, or onto your computer
- **Save an email as a PDF** (handy for filing invoices and receipts)
- **Draft emails** with attachments, replies and forwards that keep the thread, and your signature
- **Send** drafts, one or many, only after you confirm the recipients
- **Move, flag and categorise** mail, and create mail folders
- **Read, create, upload, download, move, rename and copy** files in OneDrive and SharePoint,
  including large files
- **Read and edit Excel workbooks** in place, with an automatic backup copy before each change
- **Add and change calendar events**, including all-day events and free/busy status, on your own
  calendar without inviting anyone

## What it won't do

- **Delete mail.** There are no tools that delete email, and it won't move mail to Deleted Items or
  Junk.
- **Delete files without asking.** `delete_files` sends files and empty folders (never folders with anything in them) to the recycle bin, and
  you confirm every deletion. Turn it off with `m365-mcp configure --delete-mode disabled`.
- **Invite people or delete events.** Calendar tools never add attendees, won't change events that
  have them, and can't delete events.
- **Overwrite files silently.** Name clashes get a new name unless you ask to overwrite, and
  OneDrive/SharePoint keep the old version in version history.
- **Send without asking.** Email is drafted by default. When you ask Claude to send, you're shown
  who it's going to and asked to confirm. Administrators can turn sending off entirely.
- **Phone home.** There's no server, telemetry or third party. Your mail and files go between
  Microsoft and your computer, and to the AI app you're using when it reads them.

## Setup

### 1. Register the app in Entra (once per organisation)

An IT administrator (or anyone allowed to register apps) creates an app registration in your
organisation's Entra ID. It takes a couple of minutes with the included script, by hand, or by
asking Claude in Chrome to do it. See **[docs/entra-setup.md](docs/entra-setup.md)**.

You'll end up with a **Client ID** and a **Tenant ID**. They aren't secret, but they're specific to
your organisation, so keep them out of public places like this repository.

### 2. Install (each person)

Install [uv](https://docs.astral.sh/uv/getting-started/installation/), then:

```bash
uv tool install git+https://github.com/dcohen410/m365-mcp
m365-mcp configure --client-id <client-id> --tenant-id <tenant-id>
m365-mcp login
```

`login` opens your browser to sign in to Microsoft 365. Your sign-in is stored in the macOS
Keychain or Windows Credential Manager, not in a plain file. On a machine without a browser, use
`m365-mcp login --device-code`.

### 3. Connect it to Claude

**Claude Code**

```bash
claude mcp add m365 -- m365-mcp serve
```

**Claude Desktop**: Settings → Developer → Edit Config, then add the following to
`claude_desktop_config.json` and restart Claude. Run `which m365-mcp` (macOS) or
`where m365-mcp` (Windows) to get the full path.

```json
{
  "mcpServers": {
    "m365": {
      "command": "/Users/you/.local/bin/m365-mcp",
      "args": ["serve"]
    }
  }
}
```

Claude asks your permission before using each tool. You can choose "Always allow" for the
read-only ones, but leave `send_drafts` and `delete_files` on "Ask every time".

## Options

```bash
m365-mcp configure --local-dir ~/Downloads --local-dir ~/Documents/Invoices   # folders it may use on this computer
m365-mcp configure --signature ~/signature.html                               # added to new drafts
m365-mcp configure --signature-image ~/signature.png                          # embedded at the end of new drafts
m365-mcp configure --send-mode disabled                                        # drafts only, never send
m365-mcp configure --read-only                                                 # look but don't touch (--no-read-only to undo)
m365-mcp configure --tools mail,drafts                                         # only offer some groups of tools
m365-mcp status                                                                # show settings and who's signed in
m365-mcp logout                                                                # sign out and delete the saved sign-in
```

The same settings can come from environment variables: `M365_MCP_CLIENT_ID`, `M365_MCP_TENANT_ID`,
`M365_MCP_LOCAL_DIRS` (separated by `:` on macOS/Linux, `;` on Windows), `M365_MCP_SIGNATURE_HTML`,
`M365_MCP_SIGNATURE_IMAGE`, `M365_MCP_SEND_MODE`, `M365_MCP_READ_ONLY` (`true`/`false`) and
`M365_MCP_TOOLS` (e.g. `mail,files`).

**Signature image**: a PNG, JPEG or GIF (up to 1 MB) embedded in the email itself, not linked, so it
shows without the recipient having to download pictures. Images wider than 700 pixels are shown at
700. `configure` keeps its own copy in the settings folder.

**Read-only mode** offers only the tools that search and read: nothing can be moved, written,
drafted or sent. **Tool groups** are `mail`, `drafts` (drafting and sending), `files` (OneDrive and
SharePoint), `excel` and `calendar`; the default is all of them. Tools that are switched off aren't shown to
Claude at all. Restart Claude after changing either setting.

Logs never contain sign-in tokens or email addresses; they're replaced with `[redacted]` and
`[email]`, so a log is safe to share when reporting a problem.

By default it may only read and write files on your computer inside `~/Downloads`. It doesn't need
local access for OneDrive or SharePoint work.

## Things to ask Claude

- "Find Alex's email with the signed contract and save the PDF to HR/Contracts/Alex in the Team SharePoint site."
- "Save the supplier invoice email as a PDF in Finance/Invoices/2026, then archive it."
- "Draft a reply to Sam's last email saying Thursday works, and attach the agenda from my OneDrive."
- "Add a row to the Team Register spreadsheet for our new starter."
- "Send the three drafts we just wrote."
- "Put Sam's leave from 9 to 26 November in my calendar as an all-day event, shown as free."

## Tools

| Area | Tools |
|---|---|
| Mail | `search_mail`, `read_email`, `list_mail_folders`, `create_mail_folder`, `save_email_attachments`, `save_email`, `move_emails`, `mark_emails`, `create_draft`, `send_drafts` |
| Files | `find_sharepoint_sites`, `list_libraries`, `list_folder`, `search_files`, `read_file`, `download_file`, `upload_file`, `write_text_file`, `create_folder`, `move_or_rename`, `copy_file`, `delete_files` |
| Excel | `excel_list_sheets`, `excel_read`, `excel_write`, `excel_append_rows` |
| Calendar | `list_events`, `create_event`, `update_event` |

A note on safety: emails are written by other people, and a message could contain text that tries to
instruct an AI. The server labels email content as untrusted, and sending always needs your
confirmation, but stay alert when you ask Claude to act on what an email says.

## Development

```bash
git clone https://github.com/dcohen410/m365-mcp && cd m365-mcp
uv venv && uv pip install -e '.[dev]'
uv run pytest
```

Tests mock Microsoft Graph, so they need no account. Please never commit real email content, client
IDs, tenant IDs or tokens.

## Licence

MIT. See [LICENSE](LICENSE).
