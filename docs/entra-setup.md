# Setting up the Entra app registration

m365-mcp signs in with an app registration that lives in **your own** Microsoft 365 organisation.
Someone with permission to register apps does this once; everyone else just uses the two IDs it
produces. There is no client secret.

There are three ways to do it. Pick one.

## Option A: the setup script (quickest)

Needs the [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli).

```bash
az login --allow-no-subscriptions
./scripts/setup-entra.sh            # macOS / Linux
.\scripts\setup-entra.ps1           # Windows PowerShell
```

It prints the **Client ID** and **Tenant ID**, and grants admin consent if you're an admin.

## Option B: by hand in the Entra admin centre

1. Go to [entra.microsoft.com](https://entra.microsoft.com) → **Identity → Applications → App registrations → New registration**.
2. Name: `m365-mcp`. Supported account types: **Accounts in this organizational directory only**.
3. Redirect URI: platform **Public client/native (mobile & desktop)**, value `http://localhost`. Click **Register**.
4. On the app's **Authentication** page, set **Allow public client flows** to **Yes** and save.
   This is easy to miss, and sign-in fails without it.
5. On **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
6. Click **Grant admin consent for <your organisation>** and confirm. (If you're not an admin, send
   an admin this link: `https://login.microsoftonline.com/<tenant-id>/adminconsent?client_id=<client-id>`.)
7. From the app's **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID**.

Don't add **Application** permissions and don't create a client secret. m365-mcp only needs
delegated permissions, which means it can only ever do what the signed-in person can already do.

## Option C: let Claude do it in your browser (Claude in Chrome)

If you use Claude with the Claude in Chrome extension, sign in to
[entra.microsoft.com](https://entra.microsoft.com) as an administrator in Chrome, then paste this to
Claude:

> Please set up an Entra app registration for m365-mcp in the Entra admin centre tab I'm signed in to.
> Follow Option B in https://github.com/dcohen410/m365-mcp/blob/main/docs/entra-setup.md exactly:
> single-tenant, public client redirect URI `http://localhost`, "Allow public client flows" set to Yes,
> and only these delegated Microsoft Graph permissions: User.Read, Mail.ReadWrite, Mail.Send,
> Files.ReadWrite.All, Sites.ReadWrite.All. Don't create a client secret or add application
> permissions. Before clicking "Grant admin consent", show me the permission list and wait for my OK.
> When you're done, tell me the Application (client) ID and Directory (tenant) ID.

## Why these permissions

| Permission | Used for |
|---|---|
| `User.Read` | Knowing who is signed in |
| `Mail.ReadWrite` | Searching and reading mail, saving attachments, creating drafts, moving and flagging mail |
| `Mail.Send` | Sending a draft, only after the user confirms |
| `Files.ReadWrite.All` | Reading and writing files in OneDrive and in libraries shared with you |
| `Sites.ReadWrite.All` | Finding SharePoint sites and working in their document libraries |

All are delegated: Microsoft only lets the app reach mailboxes, files and sites the signed-in person
already has access to.

## Restricting who can use it (optional)

On the app's **Enterprise applications** entry, under **Properties**, set **Assignment required?**
to **Yes**, then add the people or groups allowed to sign in under **Users and groups**.
