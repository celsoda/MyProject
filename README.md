# Outlook Connector (Claude Desktop ↔ Microsoft Graph)

**LEFT:** [Claude Desktop](https://claude.ai/download) via **MCP (stdio)**  
**RIGHT:** Outlook / Microsoft Graph with **delegated** user OAuth (device code, public client — **no client secret**)

Claude calls tools like `list_mail` / `send_mail`. This process holds Microsoft tokens in an **encrypted local cache**. Claude never receives Graph access tokens.

```
Claude Desktop  --stdio MCP-->  outlook-connector  --delegated Graph-->  Outlook
   (tools)                      (this repo)           Mail.Read/Send
```

An optional HTTP API (`npm start`) shares the same connector core for curl/debugging. Prefer MCP for Claude Desktop.

---

## Quick start (mock smoke, no Azure)

```bash
npm install
npm run smoke
```

Expected: JSON with `"ok": true` and mock inbox subjects.

---

## Live test with Claude Desktop

### 1) Create an Entra (Azure AD) app — public client

Do this in [Microsoft Entra admin center](https://entra.microsoft.com/) or [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade):

1. **New registration**
   - Name: e.g. `outlook-connector-local`
   - Supported account types:
     - **Personal Microsoft accounts only** → later set `AZURE_TENANT=consumers`, or
     - **Accounts in any org + personal** → `AZURE_TENANT=common` (good first try)
   - Redirect URI: skip for device code, **or** add platform **Mobile and desktop applications** with  
     `https://login.microsoftonline.com/common/oauth2/nativeclient` (optional)
2. After create, copy **Application (client) ID** → this is `AZURE_CLIENT_ID`.
3. **Authentication**
   - Under **Advanced settings**, set **Allow public client flows** = **Yes** (required for device code).
   - Do **not** create a client secret for this spike.
4. **API permissions** → **Microsoft Graph** → **Delegated** (not Application):
   - `User.Read`
   - `Mail.Read`
   - `Mail.Send`
   - (`offline_access` is requested by the app at runtime; you do not always add it in the portal UI)
   - Click **Add permissions**. Do **not** click “Grant admin consent” unless IT is deliberately approving the app for everyone.
5. User consent: the first login will prompt **you** (or fail if the company disabled user consent — see the architecture plan).

**Do not paste client secrets into chat.** This spike uses a public client + device code only.

### 2) Configure env on your machine

```bash
cp .env.example .env
# edit .env — set AZURE_CLIENT_ID and AZURE_TENANT=common|consumers
export $(grep -v '^#' .env | xargs)   # or use your usual env loader
npm install
```

Optional CLI login (prints the device code in the terminal):

```bash
npm run login
```

Open the URL, enter the code, sign in to Microsoft, approve mail scopes. Tokens land in `.data/msal-cache.enc`.

### 3) Point Claude Desktop at this MCP server

Edit Claude Desktop config:

- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook-connector": {
      "command": "npm",
      "args": ["run", "mcp"],
      "cwd": "/ABSOLUTE/PATH/TO/MyProject",
      "env": {
        "AZURE_CLIENT_ID": "YOUR-APPLICATION-CLIENT-ID",
        "AZURE_TENANT": "common",
        "DATA_DIR": "/ABSOLUTE/PATH/TO/MyProject/.data"
      }
    }
  }
}
```

Notes:

- Replace `/ABSOLUTE/PATH/TO/MyProject` with your clone path.
- Use `AZURE_TENANT=consumers` for personal `@outlook.com` / MSA-only tests.
- Leave `MOCK_GRAPH` unset for live Outlook.
- Fully quit and reopen Claude Desktop after saving.

If `cwd` is ignored on your Claude Desktop build, use:

```json
{
  "mcpServers": {
    "outlook-connector": {
      "command": "npx",
      "args": [
        "--yes",
        "tsx",
        "/ABSOLUTE/PATH/TO/MyProject/src/mcpServer.ts"
      ],
      "env": {
        "AZURE_CLIENT_ID": "YOUR-APPLICATION-CLIENT-ID",
        "AZURE_TENANT": "common",
        "DATA_DIR": "/ABSOLUTE/PATH/TO/MyProject/.data"
      }
    }
  }
}
```

### 4) Exercise tools in Claude Desktop

1. Confirm the `outlook-connector` MCP server shows as connected.
2. Ask Claude to call **`outlook_login`** → open the verification URL, enter the user code, consent.
3. Ask Claude to call **`list_mail`** (or “list my recent Outlook inbox mail via the connector”).
4. Optional: **`send_mail`** to an address you control.
5. **`auth_status`** / **`outlook_logout`** as needed.

---

## MCP tools

| Tool | Purpose |
| --- | --- |
| `auth_status` | Linked? username? mock vs live |
| `outlook_login` | Start device-code login; returns `userCode` + `verificationUri` |
| `outlook_logout` | Wipe local encrypted MSAL cache |
| `list_mail` | Recent Inbox messages (`limit` optional) |
| `send_mail` | Send mail as the linked user (`to`, `subject`, `bodyText`) |

---

## Optional HTTP API (same connector)

```bash
npm start
# POST /v1/auth/device-code/start
# POST /v1/auth/device-code/wait
# GET  /v1/mail/messages?limit=10
# POST /v1/mail/messages  { "to":["you@example.com"], "subject":"hi", "bodyText":"..." }
# GET  /v1/auth/status
# DELETE /v1/auth
```

Set `CONNECTOR_API_KEY` to require header `x-api-key` on `/v1/*`.

---

## Cursor / API clients

- **Cursor:** add the same MCP server command in Cursor MCP settings (stdio), mirroring the Claude Desktop snippet.
- **HTTP:** any agent/tool-caller can use `/v1/*` instead of MCP; still no Microsoft tokens on the left.

---

## Security (spike)

- Public client + device code → **no client secret** in env or chat.
- MSAL token cache encrypted at rest under `DATA_DIR` (AES-256-GCM); key auto-created as `.data/local-encryption.key` unless `TOKEN_ENCRYPTION_KEY` is set.
- Delegated scopes only; no app-only Graph permissions.
- Company tenants that disable user consent will still show “Need admin approval” — that is an IT policy limit, not a bug in this spike.

---

## Scripts

| Script | Meaning |
| --- | --- |
| `npm run smoke` | Mock end-to-end without Azure |
| `npm run login` | Interactive device-code login |
| `npm run mcp` | Claude Desktop stdio MCP server |
| `npm start` | Optional HTTP server on `PORT` (default 8787) |
| `npm run typecheck` | `tsc --noEmit` |
