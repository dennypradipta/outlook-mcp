# Multi-User Support for the Outlook MCP Server — Design

- **Date:** 2026-10-20
- **Status:** Approved for planning
- **Author:** (agent) with Denny
- **Topic:** make the MCP server serve many StartGPT users, each with their own
  Microsoft account, instead of a single shared token file.

## 1. Problem

Today the server is single-user:

- One token file, `$HOME/.outlook-mcp-tokens.json`, is shared by every request.
- The Streamable HTTP endpoint `/mcp` has no notion of identity. Whoever can reach
  the port uses the one and only token.
- The OAuth callback always writes that same file.

StartGPT is multi-user. Each request arrives with the identity of the user it is
acting on behalf of. The server must use **that user's** Microsoft mailbox, calendar,
OneDrive, and Power Automate, with their own consent and their own tokens.

## 2. Goals

- Identify the current user per request and use only that user's Microsoft tokens.
- Store and refresh Microsoft tokens **per user**, isolated from one another.
- Let each user complete Microsoft consent once, via the existing `authenticate` tool.
- Keep the endpoint stateless per request (no server-side sessions).
- Change as little of the ~34 existing tool handlers as possible.
- Preserve the current single-user behavior for stdio (local) usage.

## 3. Non-goals

- No MCP-level OAuth / resource-server implementation (the MCP server does not
  issue its own tokens).
- No admin UI, no user provisioning, no role management.
- No database. Token storage stays file-based (swappable later behind the same
  interface).
- No change to which Microsoft Graph scopes are requested.

## 4. Key decisions

| Decision | Choice | Rationale |
|---|---|---|
| User identity source | HTTP header `x-avenu-lh-user-email` (configurable) | Provided by the trusted StartGPT gateway. |
| Trust model | Header is trusted as-is | Confirmed: `/mcp` is only reachable through the StartGPT gateway, which strips any client-supplied value. No extra signature/verification. |
| Identity value | User email | This is what the gateway sends. Emails can change; a stable id would be better but is not available today. Documented as a known limitation. |
| Token storage | One file per user, named by `sha256(email)` | Safe filenames, no collisions, trivial on a shared Docker volume. |
| Identity propagation | `AsyncLocalStorage` | Tool handlers keep calling `ensureAuthenticated()` unchanged; stdio keeps working when the context is empty. |
| Consent linking | Server-issued nonce stored on disk, passed as `?n=<nonce>` | Prevents a user from starting a login "as" another user's email. |

## 5. Architecture

```
StartGPT -- POST /mcp (x-avenu-lh-user-email: alice@x.com) -->
  index.js: read header -> AsyncLocalStorage(userEmail) -> transport.handleRequest
    tool handler -> ensureAuthenticated()
        -> TokenStorage for that user -> Microsoft Graph
        -> token expired -> refresh with that user's refresh_token -> save
    no token -> throw Error('Authentication required')
        -> tool returns "use the authenticate tool"
        -> authenticate tool: write pending nonce -> return /auth?n=<nonce>
        -> user opens URL -> auth server validates nonce -> Microsoft login
        -> callback: write token file for that email
```

## 6. Components

### 6.1 Request context — `utils/request-context.js` (new)

- Wraps Node's `AsyncLocalStorage` holding `{ userEmail }`.
- `runWithUser(userEmail, fn)` — run `fn` with the user bound.
- `getCurrentUser()` — returns the current email, or `undefined` outside a bound
  context (e.g. stdio).
- Header name comes from `MCP_USER_HEADER` (default `x-avenu-lh-user-email`);
  header lookup is case-insensitive (Node lowercases header keys).

### 6.2 Per-user paths — `auth/user-paths.js` (new)

Single source of truth for where tokens and pending logins live, shared by the MCP
server and the auth server (same image, same volume).

- `getCurrentUser()` — re-exported from `utils/request-context`.
- `getStoreRoot()` — `OUTLOOK_TOKEN_STORE_DIR`, default `$HOME`.
- `tokenDir()` — `<root>/tokens`.
- `tokenPathForUser(email)` — `<root>/tokens/<sha256(email)>.json`; when `email` is
  falsy, returns the legacy global path `$HOME/.outlook-mcp-tokens.json`.
- `pendingDir()` — `<root>/pending`.
- `pendingLoginPath(nonce)` — `<root>/pending/<nonce>.json`.
- `userKey(email)` — `sha256(email.toLowerCase().trim())` hex.

### 6.3 Token manager — `auth/token-manager.js` (modified)

- Replace every use of `config.AUTH_CONFIG.tokenStorePath` with
  `tokenPathForUser(getCurrentUser())`.
- Affected functions: `loadTokenCache`, `saveTokenCache`, `getAccessToken`,
  `getFlowAccessToken`, `saveFlowTokens`, `createTestTokens`.
- Signatures stay the same; they become user-aware transparently. Sync functions
  (`getFlowAccessToken`) remain sync (they read that user's file).

### 6.4 Auth entry — `auth/index.js` (modified)

- `ensureAuthenticated(forceNew)`:
  - Resolve `email = getCurrentUser()`.
  - Get a cached `TokenStorage` for that user from a module-level
    `Map<key, TokenStorage>` (key = `userKey(email)` or `'__default__'`), created as
    `new TokenStorage({ tokenStorePath: tokenPathForUser(email) })`.
  - `await store.getValidAccessToken()`; if falsy, `throw new Error('Authentication required')`.
- The thrown message stays exactly `'Authentication required'` so existing tool
  `catch` blocks and unit tests remain valid.
- Export `getTokenStorageForCurrentUser()` for reuse by auth tools.

### 6.5 Auth tools — `auth/tools.js` (modified)

- `handleAuthenticate(args)`:
  - Test mode: `tokenManager.createTestTokens()` (now writes the current user's file)
    and return the existing success text.
  - Real: resolve `email = getCurrentUser()`.
    - If `email`: generate `nonce = crypto.randomBytes(24).toString('hex')`, write
      `pendingLoginPath(nonce)` = `{ email, createdAt }`, return
      `${authServerUrl}/auth?n=${nonce}`.
    - If no `email` (stdio): return `${authServerUrl}/auth` (legacy single-user).
- `handleCheckAuthStatus()`: read the current user's tokens via the per-user store.
- `handleAbout()`: unchanged.

### 6.6 Server wiring — `index.js` (modified)

- In the HTTP handler, before `transport.handleRequest(req, res)`:
  - `const email = req.headers[process.env.MCP_USER_HEADER || 'x-avenu-lh-user-email'];`
  - `await runWithUser(email, () => transport.handleRequest(req, res));`
- `/health`, `/healthz`, CORS, and 405 handling unchanged.
- Stdio path unchanged (no ALS context, so tokens fall back to the global file).

### 6.7 Auth server — `outlook-auth-server.js` (modified)

- `/auth`:
  - If `?n=<nonce>` present: read `pendingLoginPath(nonce)`, verify it exists and is
    within TTL (10 min). Resolve `email`. If invalid/expired → 403 page.
  - Else: legacy single-user (global file).
  - Generate `state`, store `state -> { email }` (in-memory `Map`, TTL cleanup as
    today), redirect to Microsoft.
- `/auth/callback`:
  - Resolve `email` from `state`.
  - Exchange code for tokens, then write to `tokenPathForUser(email)` with mode `0600`.
  - Delete the consumed nonce file.
- `exchangeCodeForTokens` takes the destination path (or the email) instead of the
  hardcoded `AUTH_CONFIG.tokenStorePath`.

## 7. Data flow (first use by a new user)

1. StartGPT sends a tool call to `/mcp` with `x-avenu-lh-user-email: alice@x.com`.
2. `index.js` binds the email and runs the transport.
3. The tool calls `ensureAuthenticated()` → no token file for Alice → throws
   `Authentication required` → the tool replies "use the authenticate tool".
4. The agent calls `authenticate` → server writes `pending/<nonce>.json` and returns
   `.../auth?n=<nonce>`.
5. Alice opens the URL → auth server validates the nonce → Microsoft login → callback
   writes `tokens/<sha256(alice)>.json`.
6. Subsequent tool calls for Alice succeed; expired tokens are refreshed with Alice's
   refresh token and re-saved.

## 8. Configuration

| Env | Default | Purpose |
|---|---|---|
| `MCP_USER_HEADER` | `x-avenu-lh-user-email` | header carrying the user's identity |
| `OUTLOOK_TOKEN_STORE_DIR` | `$HOME` | root for `tokens/` and `pending/` |
| `MCP_AUTH_SERVER_URL` | `http://localhost:3333` | base URL returned by `authenticate` |
| `MS_REDIRECT_URI` | `http://localhost:3333/auth/callback` | OAuth redirect (must match Azure) |

## 9. Docker

- Both services (`mcp` and `auth`) set `OUTLOOK_TOKEN_STORE_DIR=/data` and keep the
  shared `mcp-tokens` volume at `/data`. Tokens live in `/data/tokens/`, pending logins
  in `/data/pending/`.
- Production dependency (outside the code): because the OAuth callback is opened in the
  **user's browser** and reached by **Microsoft's redirect**, the auth service must be
  exposed at a public URL. `MS_REDIRECT_URI` and `MCP_AUTH_SERVER_URL` must be set to
  that public URL, not `localhost`.

## 10. Backward compatibility

- stdio (Claude Desktop) and any HTTP request without the identity header use the
  legacy global token file and behave exactly as before.
- `ensureAuthenticated()`'s error message is unchanged, so existing tool error
  handling and unit tests keep passing.

## 11. Testing

- **Unit**
  - `auth/user-paths`: `userKey` normalization, `tokenPathForUser` (per-user + legacy
    fallback), `pendingLoginPath`.
  - Pending nonce: valid, expired, tampered/unknown.
  - `request-context`: `runWithUser` isolation across interleaved async work.
- **Integration** (mock Microsoft Graph)
  - Two users with distinct token files stay isolated; tool A never reads B's token.
  - HTTP request without header falls back to the legacy file.
  - Refresh path writes back to the correct user's file.
- **Regression**: existing Jest suites (they mock `ensureAuthenticated`) stay green.
- **Manual**: `docker run` with an identity header, exercise authenticate + a tool in
  both test mode and real mode.

## 12. Security considerations

- The identity header is trusted because the gateway is the only path to `/mcp` and
  strips client-supplied values. If that assumption ever breaks, the whole model is
  impersonable — this must be recorded and re-checked if the network topology changes.
- Refresh tokens for every user live on the shared volume. The volume must be
  access-controlled.
- Consent is bound to a short-lived server-issued nonce to prevent account mix-up.

## 13. Out of scope / future

- Stable user id instead of email.
- Redis/DB token store for multi-replica deployments (the file layout is isolated
  behind `auth/user-paths.js` + `TokenStorage`).
- MCP-level OAuth if clients ever connect without the StartGPT gateway.
