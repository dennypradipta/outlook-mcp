[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# M365 Assistant MCP Server

A comprehensive MCP (Model Context Protocol) server that connects Claude with Microsoft 365 services through the Microsoft Graph API and Power Automate API.

## Supported Services

- **Outlook** - Email, calendar, folders, and rules
- **OneDrive** - Files, folders, search, and sharing
- **Power Automate** - Flows, environments, and run history

## Directory Structure

```
├── index.js                 # Main entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-manager.js     # Token storage and refresh (Graph + Flow)
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── delete.js            # Delete event
│   ├── cancel.js            # Cancel event
│   ├── accept.js            # Accept event
│   └── decline.js           # Decline event
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   ├── send.js              # Send email
│   └── mark-as-read.js      # Mark email read/unread
├── folder/                  # Folder functionality
│   ├── index.js             # Folder exports
│   ├── list.js              # List folders
│   ├── create.js            # Create folder
│   └── move.js              # Move emails
├── rules/                   # Email rules functionality
│   ├── index.js             # Rules exports
│   ├── list.js              # List rules
│   └── create.js            # Create rule
├── onedrive/                # OneDrive functionality
│   ├── index.js             # OneDrive exports
│   ├── list.js              # List files/folders
│   ├── search.js            # Search files
│   ├── download.js          # Get download URL
│   ├── upload.js            # Simple upload (<4MB)
│   ├── upload-large.js      # Chunked upload (>4MB)
│   ├── share.js             # Create sharing link
│   └── folder.js            # Create/delete folders
├── power-automate/          # Power Automate functionality
│   ├── index.js             # Power Automate exports
│   ├── flow-api.js          # Flow API client
│   ├── list-environments.js # List environments
│   ├── list-flows.js        # List flows
│   ├── run-flow.js          # Trigger flow
│   ├── list-runs.js         # Run history
│   └── toggle-flow.js       # Enable/disable flow
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: OAuth 2.0 authentication with Microsoft Graph API (+ Flow API for Power Automate)
- **Email Management**: List, search, read, send, and organize emails
- **Calendar Management**: List, create, accept, decline, and delete calendar events
- **OneDrive Integration**: List, search, upload, download, and share files
- **Power Automate**: List environments/flows, trigger flows, view run history
- **Modular Structure**: Clean separation of concerns for maintainability
- **Test Mode**: Simulated responses for testing without real API calls

## Available Tools

### Outlook (Email & Calendar)
| Tool | Description |
|------|-------------|
| `list-emails` | List recent emails from inbox |
| `search-emails` | Search emails with filters |
| `read-email` | Read email content |
| `send-email` | Send a new email |
| `mark-as-read` | Mark email as read/unread |
| `list-events` | List calendar events |
| `create-event` | Create calendar event |
| `accept-event` | Accept event invitation |
| `decline-event` | Decline event invitation |
| `delete-event` | Delete calendar event |
| `list-folders` | List mail folders |
| `create-folder` | Create mail folder |
| `move-emails` | Move emails between folders |
| `list-rules` | List inbox rules |
| `create-rule` | Create inbox rule |

### OneDrive
| Tool | Description |
|------|-------------|
| `onedrive-list` | List files in a path |
| `onedrive-search` | Search files by query |
| `onedrive-download` | Get download URL |
| `onedrive-upload` | Upload small file (<4MB) |
| `onedrive-upload-large` | Chunked upload (>4MB) |
| `onedrive-share` | Create sharing link |
| `onedrive-create-folder` | Create folder |
| `onedrive-delete` | Delete file or folder |

### Power Automate
| Tool | Description |
|------|-------------|
| `flow-list-environments` | List Power Platform environments |
| `flow-list` | List flows in environment |
| `flow-run` | Trigger a manual flow |
| `flow-list-runs` | Get flow run history |
| `flow-toggle` | Enable/disable a flow |

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register app in Azure Portal (see detailed steps below)
3. **Configure environment**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Configure Claude**: Update your Claude Desktop config with the server path
5. **Start auth server**: `npm run auth-server`
6. **Authenticate**: Use the authenticate tool in Claude to get the OAuth URL
7. **Start using**: Access your M365 data through Claude!

## Installation

### Prerequisites
- Node.js 14.0.0 or higher
- npm or yarn package manager
- Azure account for app registration

### Install Dependencies

```bash
npm install
```

## Docker

The server ships with a `Dockerfile` and a `docker-compose.yml` that run it over the
**Streamable HTTP** transport (the container is a long-lived HTTP service, not a
stdio process). Two services share one image and one token volume:

- **`mcp`** — the MCP server on `http://localhost:3000/mcp` (health check at `/health`)
- **`auth`** — the OAuth callback server on `http://localhost:3333`

### Quick start

The `docker-compose.yml` runs the prebuilt image
`hyperjump.azurecr.io/startgpt/mcp-outlook:latest` — no local build required.

```bash
# 1. Log in to Azure Container Registry (once)
az acr login --name hyperjump        # or: docker login hyperjump.azurecr.io

# 2. Provide credentials
cp .env.example .env                 # add MS_CLIENT_ID / MS_CLIENT_SECRET / MS_TENANT_ID

# 3. Run
docker compose up -d
```

Then authenticate once (tokens are stored in the shared `mcp-tokens` volume):

1. Start the auth flow — either open `http://localhost:3333/auth` in a browser, or call
   the `authenticate` tool through an MCP client and open the URL it returns.
2. Sign in with your Microsoft account. The callback writes tokens to `/data` in the
   shared volume, which the `mcp` service reads.

### Connecting an MCP client

Point any Streamable HTTP MCP client at `http://localhost:3000/mcp`. For example, with
Claude Code:

```bash
claude mcp add --transport http m365 http://localhost:3000/mcp
```

or in a JSON config:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "type": "http",
      "url": "http://localhost:3000/mcp"
    }
  }
}
```

### Running without Compose

The same image can be run directly:

```bash
az acr login --name hyperjump        # or: docker login hyperjump.azurecr.io

# MCP server (Streamable HTTP on :3000, token volume mounted at /data)
docker run -d --name outlook-mcp \
  -p 3000:3000 \
  --env-file .env \
  -v mcp-tokens:/data \
  hyperjump.azurecr.io/startgpt/mcp-outlook:latest

# OAuth callback server (run when you need to (re)authenticate)
docker run --rm -it \
  -p 3333:3333 \
  --env-file .env \
  -v mcp-tokens:/data \
  hyperjump.azurecr.io/startgpt/mcp-outlook:latest node outlook-auth-server.js
```

To use the image with a **stdio** MCP client instead, override the transport and keep
stdin attached:

```bash
docker run -i --rm \
  -e MCP_TRANSPORT=stdio \
  --env-file .env \
  -v mcp-tokens:/data \
  hyperjump.azurecr.io/startgpt/mcp-outlook:latest
```

### Docker configuration

| Variable | Default | Used by | Description |
|----------|---------|---------|-------------|
| `MCP_TRANSPORT` | `stdio` | mcp | `stdio` or `http` (Streamable HTTP) |
| `PORT` / `MCP_PORT` | `3000` | mcp | HTTP listen port |
| `HOST` / `MCP_HOST` | `0.0.0.0` | mcp | HTTP bind address |
| `MCP_PATH` | `/mcp` | mcp | HTTP endpoint path |
| `MCP_AUTH_SERVER_URL` | `http://localhost:3333` | mcp | URL returned by the `authenticate` tool |
| `AUTH_PORT` | `3333` | auth | OAuth callback server port |
| `MS_REDIRECT_URI` | `http://localhost:3333/auth/callback` | auth | Must match the Azure app registration |
| `OUTLOOK_TOKEN_STORE_DIR` | `$HOME` | both | Root for per-user `tokens/` and `pending/` |
| `MCP_USER_HEADER` | `x-avenu-lh-user-email` | mcp | Trusted identity header (multi-user) |

> The redirect URI is baked into the browser-facing OAuth flow. If the auth service is
> not reachable at `localhost:3333` from your browser, set `MS_REDIRECT_URI` and
> `MCP_AUTH_SERVER_URL` accordingly and register the matching redirect URI in Azure.

### Multi-user

When deployed behind the StartGPT gateway, every request carries the acting user in a
trusted header (default `x-avenu-lh-user-email`). The server stores and refreshes a
separate Microsoft token set per user under `$OUTLOOK_TOKEN_STORE_DIR/tokens/` (a file
per user, named `sha256(email).json`). Each user authenticates once through the
`authenticate` tool; the returned URL is scoped to that user.

- Identity header is configurable via `MCP_USER_HEADER`.
- Token store root is configurable via `OUTLOOK_TOKEN_STORE_DIR` (Docker: `/data`).
- Without the gateway/header (e.g. local stdio), behavior falls back to the single
  legacy token file `~/.outlook-mcp-tokens.json`.

Production note: the OAuth callback is opened in the user's browser and reached by
Microsoft's redirect, so the auth service must be published at a public URL, and
`MS_REDIRECT_URI` / `MCP_AUTH_SERVER_URL` must point at it.

## Azure App Registration & Configuration

### App Registration

1. Open [Azure Portal](https://portal.azure.com/)
2. Search for "App registrations"
3. Click "New registration"
4. Name: "M365 MCP Server"
5. Account type: "Accounts in any organizational directory and personal Microsoft accounts"
6. Redirect URI: Web → `http://localhost:3333/auth/callback`
7. Click "Register"
8. Copy the "Application (client) ID" for your `.env` file

### App Permissions

1. Go to "API permissions" under Manage
2. Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
3. Add these permissions:
   - `offline_access`
   - `User.Read`
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
4. Click "Add permissions"

**For Power Automate** (optional):
- Requires additional Azure AD configuration with Flow API scope
- See Power Automate section below for details

### Client Secret

1. Go to "Certificates & secrets" → "Client secrets"
2. Click "New client secret"
3. Add description and select expiration
4. **Copy the VALUE** (not the Secret ID)

## Configuration

### 1. Environment Variables

```bash
cp .env.example .env
```

Edit `.env`:
```bash
# Get these values from Azure Portal > App Registrations > Your App
MS_CLIENT_ID=your-application-client-id-here
MS_CLIENT_SECRET=your-client-secret-VALUE-here
MS_TENANT_ID=your-tenant-id-here
USE_TEST_MODE=false
```

**Important Notes:**
- Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET` in the `.env` file
- Set `MS_TENANT_ID` for single-tenant apps to avoid `/common` endpoint errors
- For Claude Desktop config, you'll use `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET`
- Always use the client secret **VALUE**, never the Secret ID

### 2. Claude Desktop Configuration

Add to your Claude Desktop config:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Authentication

### Graph API (Outlook + OneDrive)

1. Start auth server: `npm run auth-server`
2. Use the `authenticate` tool in Claude
3. Visit the provided URL and sign in
4. Tokens saved to `~/.outlook-mcp-tokens.json`

### Power Automate (Optional)

Power Automate requires a separate token with the Flow API scope. Configure additional Azure AD permissions for `https://service.flow.microsoft.com//.default` scope.

**Limitations:**
- Only solution-aware flows are accessible
- Only manual trigger flows can be run via API
- Requires environment ID for most operations

## Troubleshooting

### Common Issues

**"Cannot find module"**
```bash
npm install
```

**"Port 3333 in use"**
```bash
npx kill-port 3333
npm run auth-server
```

**"Invalid client secret" (AADSTS7000215)**
- Use the secret **VALUE**, not the Secret ID

**"Authentication required"**
- Delete `~/.outlook-mcp-tokens.json` and re-authenticate

## Testing

```bash
# Run with MCP Inspector
npm run inspect

# Run in test mode (mock data)
npm run test-mode

# Run Jest tests
npm test
```

## Extending the Server

1. Create new module directory
2. Implement tool handlers in separate files
3. Export tool definitions from module index
4. Import and add to `TOOLS` array in `index.js`
