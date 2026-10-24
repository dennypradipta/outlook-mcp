# Multi-User Outlook MCP — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the MCP server serve many StartGPT users, each with their own Microsoft
account, by resolving the current user from the `x-avenu-lh-user-email` header and
storing/refreshing Microsoft tokens per user.

**Architecture:** Read the identity header at the HTTP layer and bind it to an
`AsyncLocalStorage` context. Token-manager and auth entry resolve tokens for the
current user via a file-per-user store keyed by `sha256(email)`. Existing tool
handlers are untouched; the single-user/stdio path falls back to the legacy global
token file.

**Tech Stack:** Node.js (CommonJS), `@modelcontextprotocol/sdk` Streamable HTTP
transport, `async_hooks` `AsyncLocalStorage`, Jest + supertest for tests.

**Spec:** `docs/superpowers/specs/2026-10-20-multiuser-design.md`

## Global Constraints

- Node.js >= 18 (Docker image is `node:22-alpine`). No new runtime dependencies.
- CommonJS modules; match existing style (`require`, `module.exports`).
- Identity header default: `x-avenu-lh-user-email`; override via `MCP_USER_HEADER`.
  Header keys are lowercase when read from Node's `req.headers`.
- Token store root env `OUTLOOK_TOKEN_STORE_DIR` (default `$HOME`). Tokens live in
  `<root>/tokens/<sha256(email)>.json`; pending logins in `<root>/pending/<nonce>.json`.
- Nonce TTL is **10 minutes** (`10 * 60 * 1000` ms).
- Legacy fallback: when there is no user (stdio, or HTTP without the header),
  `tokenPathForUser` returns `$HOME/.outlook-mcp-tokens.json` and behavior is unchanged.
- `ensureAuthenticated()` MUST keep throwing exactly `new Error('Authentication required')`.
- **Known pre-existing failure (do NOT fix in this plan):**
  `test/auth/token-storage.test.js › should warn if client ID or secret is missing`
  expects an outdated warning string. Leave it red; do not touch `token-storage.js`.

---

## File Structure

- Create `utils/request-context.js` — AsyncLocalStorage identity context + header parsing.
- Create `auth/user-paths.js` — path layout for per-user tokens and pending logins.
- Create `auth/pending-logins.js` — create/consume login nonces.
- Modify `auth/token-manager.js` — user-aware file paths, per-user cache.
- Modify `auth/index.js` — per-user `TokenStorage` cache; `ensureAuthenticated`.
- Modify `auth/tools.js` — `authenticate` issues nonce-scoped URLs.
- Modify `index.js` — bind header user into ALS; export handler for tests.
- Modify `outlook-auth-server.js` — nonce-aware `/auth`, per-user token write.
- Create tests: `test/request-context.test.js`, `test/auth/user-paths.test.js`,
  `test/auth/pending-logins.test.js`, `test/auth/token-manager.test.js`,
  `test/auth/ensure-authenticated.test.js`, `test/auth/tools.test.js`,
  `test/index.http.test.js`, `test/auth/auth-server.test.js`.
- Modify `docker-compose.yml`, `.env.example`, `README.md`.

---

## Task 1: Request context (identity via AsyncLocalStorage)

**Files:**
- Create: `utils/request-context.js`
- Test: `test/request-context.test.js`

**Interfaces:**
- Produces:
  - `DEFAULT_USER_HEADER: string` = `'x-avenu-lh-user-email'`
  - `runWithUser(userEmail: string|undefined, fn: Function): any`
  - `getCurrentUser(): string|undefined`
  - `getUserFromHeaders(headers: object, headerName?: string): string|undefined`

- [ ] **Step 1: Write the failing test**

Create `test/request-context.test.js`:

```js
const {
  runWithUser,
  getCurrentUser,
  getUserFromHeaders,
  DEFAULT_USER_HEADER
} = require('../utils/request-context');

describe('request-context', () => {
  it('returns undefined outside a bound context', () => {
    expect(getCurrentUser()).toBeUndefined();
  });

  it('returns the bound user inside runWithUser', () => {
    const result = runWithUser('alice@x.com', () => getCurrentUser());
    expect(result).toBe('alice@x.com');
  });

  it('isolates interleaved async contexts', async () => {
    const results = {};
    await Promise.all([
      runWithUser('a@x.com', async () => {
        await new Promise((r) => setTimeout(r, 20));
        results.a = getCurrentUser();
      }),
      runWithUser('b@x.com', async () => {
        await new Promise((r) => setTimeout(r, 5));
        results.b = getCurrentUser();
      })
    ]);
    expect(results.a).toBe('a@x.com');
    expect(results.b).toBe('b@x.com');
  });

  it('extracts and trims a header value (case-insensitive key)', () => {
    expect(getUserFromHeaders({ 'x-avenu-lh-user-email': ' alice@x.com ' })).toBe('alice@x.com');
  });

  it('returns undefined for missing or blank header', () => {
    expect(getUserFromHeaders({})).toBeUndefined();
    expect(getUserFromHeaders({ 'x-avenu-lh-user-email': '   ' })).toBeUndefined();
  });

  it('handles array header values', () => {
    expect(getUserFromHeaders({ 'x-avenu-lh-user-email': ['a@x.com', 'b@x.com'] })).toBe('a@x.com');
  });

  it('honors a custom header name', () => {
    expect(getUserFromHeaders({ 'x-user': 'c@x.com' }, 'x-user')).toBe('c@x.com');
    expect(DEFAULT_USER_HEADER).toBe('x-avenu-lh-user-email');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/request-context.test.js`
Expected: FAIL — `Cannot find module '../utils/request-context'`.

- [ ] **Step 3: Write the implementation**

Create `utils/request-context.js`:

```js
/**
 * Per-request identity context for multi-user support.
 *
 * Uses AsyncLocalStorage so downstream code (token manager) can resolve the
 * current user without threading it through every function signature.
 */
const { AsyncLocalStorage } = require('async_hooks');

const storage = new AsyncLocalStorage();

const DEFAULT_USER_HEADER = 'x-avenu-lh-user-email';

/**
 * Runs fn with the given user bound to the async context.
 * @param {string|undefined} userEmail
 * @param {Function} fn
 * @returns {*} fn's return value
 */
function runWithUser(userEmail, fn) {
  return storage.run({ userEmail: userEmail || undefined }, fn);
}

/**
 * @returns {string|undefined} current user's email, or undefined outside a context
 */
function getCurrentUser() {
  const store = storage.getStore();
  return store ? store.userEmail : undefined;
}

/**
 * Extracts the user's email from request headers.
 * @param {object} headers - Node.js req.headers (lowercased keys)
 * @param {string} [headerName] - defaults to DEFAULT_USER_HEADER
 * @returns {string|undefined}
 */
function getUserFromHeaders(headers, headerName = DEFAULT_USER_HEADER) {
  if (!headers) return undefined;
  const key = (headerName || DEFAULT_USER_HEADER).toLowerCase();
  const raw = headers[key];
  const value = Array.isArray(raw) ? raw[0] : raw;
  if (value == null) return undefined;
  const trimmed = String(value).trim();
  return trimmed || undefined;
}

module.exports = {
  DEFAULT_USER_HEADER,
  runWithUser,
  getCurrentUser,
  getUserFromHeaders
};
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/request-context.test.js`
Expected: PASS (7 tests).

- [ ] **Step 5: Commit**

```bash
git add utils/request-context.js test/request-context.test.js
git commit -m "feat(auth): add per-request identity context (AsyncLocalStorage)"
```

---

## Task 2: Per-user path layout

**Files:**
- Create: `auth/user-paths.js`
- Test: `test/auth/user-paths.test.js`

**Interfaces:**
- Consumes: `getCurrentUser` from `../utils/request-context`
- Produces:
  - `getHomeDir(): string`
  - `getStoreRoot(): string`
  - `tokenDir(): string`
  - `pendingDir(): string`
  - `userKey(email: string): string` (sha256 hex of trimmed, lowercased email)
  - `tokenPathForUser(email?: string): string`
  - `pendingLoginPath(nonce: string): string`
  - `currentTokenPath(): string`

- [ ] **Step 1: Write the failing test**

Create `test/auth/user-paths.test.js`:

```js
const path = require('path');
const crypto = require('crypto');

describe('user-paths', () => {
  const originalEnv = process.env;
  let userPaths;

  beforeEach(() => {
    jest.resetModules();
    process.env = { ...originalEnv, HOME: '/mock/home' };
    delete process.env.OUTLOOK_TOKEN_STORE_DIR;
    delete process.env.USERPROFILE;
    userPaths = require('../../auth/user-paths');
  });

  afterAll(() => {
    process.env = originalEnv;
  });

  it('normalizes the user key (trim + lowercase + sha256)', () => {
    const expected = crypto.createHash('sha256').update('alice@x.com').digest('hex');
    expect(userPaths.userKey('  Alice@X.com ')).toBe(expected);
  });

  it('builds a per-user token path under <root>/tokens/', () => {
    process.env.OUTLOOK_TOKEN_STORE_DIR = '/data';
    const key = userPaths.userKey('alice@x.com');
    expect(userPaths.tokenPathForUser('alice@x.com')).toBe(path.join('/data', 'tokens', `${key}.json`));
  });

  it('defaults the root to HOME when OUTLOOK_TOKEN_STORE_DIR is unset', () => {
    expect(userPaths.getStoreRoot()).toBe('/mock/home');
  });

  it('falls back to the legacy global token file when no email', () => {
    expect(userPaths.tokenPathForUser(undefined)).toBe(path.join('/mock/home', '.outlook-mcp-tokens.json'));
    expect(userPaths.tokenPathForUser('')).toBe(path.join('/mock/home', '.outlook-mcp-tokens.json'));
  });

  it('builds pending login paths under <root>/pending/', () => {
    process.env.OUTLOOK_TOKEN_STORE_DIR = '/data';
    expect(userPaths.pendingLoginPath('abc123')).toBe(path.join('/data', 'pending', 'abc123.json'));
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/auth/user-paths.test.js`
Expected: FAIL — `Cannot find module '../../auth/user-paths'`.

- [ ] **Step 3: Write the implementation**

Create `auth/user-paths.js`:

```js
/**
 * Where per-user tokens and pending logins live. Shared by the MCP server and
 * the auth server (same image, same volume).
 */
const crypto = require('crypto');
const os = require('os');
const path = require('path');
const { getCurrentUser } = require('../utils/request-context');

function getHomeDir() {
  return process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';
}

/** Root directory holding tokens/ and pending/. */
function getStoreRoot() {
  return process.env.OUTLOOK_TOKEN_STORE_DIR || getHomeDir();
}

function tokenDir() {
  return path.join(getStoreRoot(), 'tokens');
}

function pendingDir() {
  return path.join(getStoreRoot(), 'pending');
}

/** Stable, filesystem-safe key for an email. */
function userKey(email) {
  return crypto.createHash('sha256').update(String(email).trim().toLowerCase()).digest('hex');
}

/**
 * Token file for a user. Falls back to the legacy global file when no email.
 */
function tokenPathForUser(email) {
  if (!email) {
    return path.join(getHomeDir(), '.outlook-mcp-tokens.json');
  }
  return path.join(tokenDir(), `${userKey(email)}.json`);
}

function pendingLoginPath(nonce) {
  return path.join(pendingDir(), `${nonce}.json`);
}

/** Token path for the current async request context. */
function currentTokenPath() {
  return tokenPathForUser(getCurrentUser());
}

module.exports = {
  getHomeDir,
  getStoreRoot,
  tokenDir,
  pendingDir,
  userKey,
  tokenPathForUser,
  pendingLoginPath,
  currentTokenPath
};
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/auth/user-paths.test.js`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add auth/user-paths.js test/auth/user-paths.test.js
git commit -m "feat(auth): add per-user token/pending path layout"
```

---

## Task 3: Pending login nonces

**Files:**
- Create: `auth/pending-logins.js`
- Test: `test/auth/pending-logins.test.js`

**Interfaces:**
- Consumes: `pendingLoginPath`, `pendingDir` from `./user-paths`
- Produces:
  - `TTL_MS: number` = `600000`
  - `createPendingLogin(email: string): string` (returns nonce)
  - `consumePendingLogin(nonce: string): string|null` (returns email, single-use)

- [ ] **Step 1: Write the failing test**

Create `test/auth/pending-logins.test.js`:

```js
const fs = require('fs');
const os = require('os');
const path = require('path');

describe('pending-logins', () => {
  let tmp;
  let pendingLogins;
  let userPaths;

  beforeEach(() => {
    jest.resetModules();
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pending-'));
    process.env.OUTLOOK_TOKEN_STORE_DIR = tmp;
    pendingLogins = require('../../auth/pending-logins');
    userPaths = require('../../auth/user-paths');
  });

  it('creates a nonce and consumes it to the original email', () => {
    const nonce = pendingLogins.createPendingLogin('alice@x.com');
    expect(typeof nonce).toBe('string');
    expect(nonce).toMatch(/^[a-f0-9]+$/);
    expect(fs.existsSync(userPaths.pendingLoginPath(nonce))).toBe(true);
    expect(pendingLogins.consumePendingLogin(nonce)).toBe('alice@x.com');
  });

  it('is single-use', () => {
    const nonce = pendingLogins.createPendingLogin('bob@x.com');
    expect(pendingLogins.consumePendingLogin(nonce)).toBe('bob@x.com');
    expect(pendingLogins.consumePendingLogin(nonce)).toBeNull();
  });

  it('returns null for an unknown nonce', () => {
    expect(pendingLogins.consumePendingLogin('deadbeef')).toBeNull();
  });

  it('rejects non-hex / traversal nonces', () => {
    expect(pendingLogins.consumePendingLogin('../etc/passwd')).toBeNull();
    expect(pendingLogins.consumePendingLogin('')).toBeNull();
    expect(pendingLogins.consumePendingLogin(undefined)).toBeNull();
  });

  it('returns null and deletes an expired nonce', () => {
    const nonce = pendingLogins.createPendingLogin('carol@x.com');
    const file = userPaths.pendingLoginPath(nonce);
    const rec = JSON.parse(fs.readFileSync(file, 'utf8'));
    rec.createdAt = Date.now() - (11 * 60 * 1000);
    fs.writeFileSync(file, JSON.stringify(rec));
    expect(pendingLogins.consumePendingLogin(nonce)).toBeNull();
    expect(fs.existsSync(file)).toBe(false);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/auth/pending-logins.test.js`
Expected: FAIL — `Cannot find module '../../auth/pending-logins'`.

- [ ] **Step 3: Write the implementation**

Create `auth/pending-logins.js`:

```js
/**
 * Short-lived, single-use login nonces that bind an OAuth consent flow to the
 * user who requested it. Stored on disk so the MCP server and the auth server
 * (separate processes, shared volume) can collaborate.
 */
const fs = require('fs');
const crypto = require('crypto');
const { pendingLoginPath, pendingDir } = require('./user-paths');

const TTL_MS = 10 * 60 * 1000; // 10 minutes

const NONCE_RE = /^[a-f0-9]+$/;

/**
 * Creates a pending login for an email and returns its nonce.
 * @param {string} email
 * @returns {string} nonce
 */
function createPendingLogin(email) {
  const nonce = crypto.randomBytes(24).toString('hex');
  fs.mkdirSync(pendingDir(), { recursive: true });
  fs.writeFileSync(
    pendingLoginPath(nonce),
    JSON.stringify({ email, createdAt: Date.now() }),
    { mode: 0o600 }
  );
  return nonce;
}

/**
 * Resolves the email for a nonce, or null if unknown/expired/invalid.
 * Consumes (deletes) the record so it cannot be reused.
 * @param {string} nonce
 * @returns {string|null}
 */
function consumePendingLogin(nonce) {
  if (!nonce || typeof nonce !== 'string' || !NONCE_RE.test(nonce)) {
    return null;
  }
  const file = pendingLoginPath(nonce);
  let raw;
  try {
    raw = fs.readFileSync(file, 'utf8');
  } catch (_) {
    return null;
  }
  try {
    fs.unlinkSync(file);
  } catch (_) {
    // best effort
  }
  let record;
  try {
    record = JSON.parse(raw);
  } catch (_) {
    return null;
  }
  if (!record || !record.email) return null;
  if (Date.now() - (record.createdAt || 0) > TTL_MS) return null;
  return record.email;
}

module.exports = { TTL_MS, createPendingLogin, consumePendingLogin };
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/auth/pending-logins.test.js`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add auth/pending-logins.js test/auth/pending-logins.test.js
git commit -m "feat(auth): add single-use pending login nonces"
```

---

## Task 4: Make token-manager user-aware

Current `auth/token-manager.js` reads a single global path and keeps a **module-global**
`cachedTokens` — that global cache would leak tokens across users. This task removes
the global cache and resolves the current user's path per call.

**Files:**
- Modify: `auth/token-manager.js` (whole file)
- Test: `test/auth/token-manager.test.js`

**Interfaces:**
- Consumes: `currentTokenPath` from `./user-paths`
- Produces (unchanged names/signatures):
  - `loadTokenCache(): object|null`
  - `saveTokenCache(tokens: object): boolean`
  - `getAccessToken(): string|null`
  - `getFlowAccessToken(): string|null`
  - `saveFlowTokens(flowTokens: object): boolean`
  - `createTestTokens(): object`

- [ ] **Step 1: Write the failing test**

Create `test/auth/token-manager.test.js`:

```js
const fs = require('fs');
const os = require('os');
const path = require('path');

describe('token-manager (multi-user)', () => {
  let tmp;
  let tokenManager;
  let requestContext;
  let userPaths;

  beforeEach(() => {
    jest.resetModules();
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'tokmgr-'));
    process.env.OUTLOOK_TOKEN_STORE_DIR = tmp;
    tokenManager = require('../../auth/token-manager');
    requestContext = require('../../utils/request-context');
    userPaths = require('../../auth/user-paths');
  });

  it('saves and loads tokens per user, isolated from one another', () => {
    const future = Date.now() + 3600000;
    requestContext.runWithUser('alice@x.com', () => {
      tokenManager.saveTokenCache({ access_token: 'alice-token', expires_at: future });
    });
    requestContext.runWithUser('bob@x.com', () => {
      tokenManager.saveTokenCache({ access_token: 'bob-token', expires_at: future });
    });

    expect(fs.existsSync(userPaths.tokenPathForUser('alice@x.com'))).toBe(true);
    expect(fs.existsSync(userPaths.tokenPathForUser('bob@x.com'))).toBe(true);

    const alice = requestContext.runWithUser('alice@x.com', () => tokenManager.loadTokenCache());
    const bob = requestContext.runWithUser('bob@x.com', () => tokenManager.loadTokenCache());
    expect(alice.access_token).toBe('alice-token');
    expect(bob.access_token).toBe('bob-token');
  });

  it('getAccessToken returns null for a user with no tokens', () => {
    requestContext.runWithUser('nobody@x.com', () => {
      expect(tokenManager.getAccessToken()).toBeNull();
    });
  });

  it('does not leak a token cached for one user to another', () => {
    const future = Date.now() + 3600000;
    requestContext.runWithUser('alice@x.com', () => {
      tokenManager.saveTokenCache({ access_token: 'alice-token', expires_at: future });
      expect(tokenManager.getAccessToken()).toBe('alice-token');
    });
    requestContext.runWithUser('bob@x.com', () => {
      expect(tokenManager.getAccessToken()).toBeNull();
    });
  });

  it('writes the legacy global file when there is no user context', () => {
    const future = Date.now() + 3600000;
    tokenManager.saveTokenCache({ access_token: 'global-token', expires_at: future });
    expect(fs.existsSync(path.join(tmp, '.outlook-mcp-tokens.json'))).toBe(true);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/auth/token-manager.test.js`
Expected: FAIL — the token is written to `$HOME/.outlook-mcp-tokens.json` (not under the temp root), so `expect(fs.existsSync(userPaths.tokenPathForUser('alice@x.com')))` is `false` (and later the global cache leaks).

- [ ] **Step 3: Write the implementation**

Replace the whole contents of `auth/token-manager.js`:

```js
/**
 * Token management for Microsoft Graph API authentication.
 * User-aware: resolves the current request's user via AsyncLocalStorage.
 */
const fs = require('fs');
const path = require('path');
const { currentTokenPath } = require('./user-paths');

/**
 * Loads authentication tokens for the current user from their token file.
 * @returns {object|null} - The loaded tokens or null if not available/expired
 */
function loadTokenCache() {
  try {
    const tokenPath = currentTokenPath();

    if (!fs.existsSync(tokenPath)) {
      return null;
    }

    const tokenData = fs.readFileSync(tokenPath, 'utf8');

    try {
      const tokens = JSON.parse(tokenData);

      if (!tokens.access_token) {
        return null;
      }

      const now = Date.now();
      const expiresAt = tokens.expires_at || 0;

      if (now > expiresAt) {
        return null;
      }

      return tokens;
    } catch (parseError) {
      console.error('Error parsing token file:', parseError.message);
      return null;
    }
  } catch (error) {
    console.error('Error loading token cache:', error.message);
    return null;
  }
}

/**
 * Saves authentication tokens for the current user.
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = currentTokenPath();
    fs.mkdirSync(path.dirname(tokenPath), { recursive: true });
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2), { mode: 0o600 });
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error.message);
    return false;
  }
}

/**
 * Gets the current user's Graph API access token, if valid.
 * @returns {string|null}
 */
function getAccessToken() {
  const tokens = loadTokenCache();
  return tokens ? tokens.access_token : null;
}

/**
 * Gets the current user's Flow API access token, if valid.
 * @returns {string|null}
 */
function getFlowAccessToken() {
  const tokens = loadTokenCache();
  if (!tokens) return null;

  if (tokens.flow_access_token && tokens.flow_expires_at) {
    if (Date.now() < tokens.flow_expires_at) {
      return tokens.flow_access_token;
    }
  }

  return null;
}

/**
 * Saves Flow API tokens alongside the current user's existing Graph tokens.
 * @param {object} flowTokens
 * @returns {boolean}
 */
function saveFlowTokens(flowTokens) {
  try {
    const tokenPath = currentTokenPath();

    let existingTokens = {};
    if (fs.existsSync(tokenPath)) {
      existingTokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    }

    const mergedTokens = {
      ...existingTokens,
      flow_access_token: flowTokens.access_token,
      flow_refresh_token: flowTokens.refresh_token,
      flow_expires_at: flowTokens.expires_at || (Date.now() + (flowTokens.expires_in || 3600) * 1000)
    };

    fs.mkdirSync(path.dirname(tokenPath), { recursive: true });
    fs.writeFileSync(tokenPath, JSON.stringify(mergedTokens, null, 2), { mode: 0o600 });
    console.log('Flow tokens saved successfully');
    return true;
  } catch (error) {
    console.error('Error saving Flow tokens:', error);
    return false;
  }
}

/**
 * Creates a test access token for the current user (test mode).
 * @returns {object}
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000)
  };

  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  getFlowAccessToken,
  saveFlowTokens,
  createTestTokens
};
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/auth/token-manager.test.js`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add auth/token-manager.js test/auth/token-manager.test.js
git commit -m "feat(auth): make token-manager resolve tokens per user"
```

---

## Task 5: Per-user TokenStorage + `ensureAuthenticated`

**Files:**
- Modify: `auth/index.js`
- Test: `test/auth/ensure-authenticated.test.js`

**Interfaces:**
- Consumes: `getCurrentUser` from `../utils/request-context`; `tokenPathForUser` from
  `./user-paths`; `TokenStorage` from `./token-storage`
- Produces:
  - `ensureAuthenticated(forceNew?: boolean): Promise<string>`
  - `getTokenStorageForUser(email?: string): TokenStorage`
  - (unchanged) `tokenManager`, `authTools`

- [ ] **Step 1: Write the failing test**

Create `test/auth/ensure-authenticated.test.js`:

```js
jest.mock('../../auth/token-storage');

const TokenStorage = require('../../auth/token-storage');
const { ensureAuthenticated, getTokenStorageForUser } = require('../../auth');
const { runWithUser } = require('../../utils/request-context');

describe('ensureAuthenticated (multi-user)', () => {
  it('throws "Authentication required" when no token is available', async () => {
    TokenStorage.mockImplementation(() => ({
      getValidAccessToken: jest.fn().mockResolvedValue(null)
    }));
    await expect(runWithUser('nobody@x.com', () => ensureAuthenticated()))
      .rejects.toThrow('Authentication required');
  });

  it('returns the current user\'s access token', async () => {
    TokenStorage.mockImplementation(() => ({
      getValidAccessToken: jest.fn().mockResolvedValue('token-123')
    }));
    const token = await runWithUser('alice@x.com', () => ensureAuthenticated());
    expect(token).toBe('token-123');
  });

  it('throws when forceNew is requested', async () => {
    await expect(ensureAuthenticated(true)).rejects.toThrow('Authentication required');
  });

  it('caches exactly one store per user', () => {
    const a1 = getTokenStorageForUser('alice@x.com');
    const a2 = getTokenStorageForUser('alice@x.com');
    const b1 = getTokenStorageForUser('bob@x.com');
    expect(a1).toBe(a2);
    expect(a1).not.toBe(b1);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/auth/ensure-authenticated.test.js`
Expected: FAIL — `getTokenStorageForUser is not a function` (and the store is not per-user).

- [ ] **Step 3: Write the implementation**

Replace the whole contents of `auth/index.js`:

```js
/**
 * Authentication module for Outlook MCP server.
 * Resolves tokens for the current request's user (multi-user).
 */
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');
const { getCurrentUser } = require('../utils/request-context');
const { tokenPathForUser } = require('./user-paths');

// One TokenStorage instance per user, so its refresh-dedup guard is per-user.
const stores = new Map();

function storeKeyForUser(email) {
  return email ? email.trim().toLowerCase() : '__default__';
}

/**
 * Returns the cached TokenStorage for a user (or the legacy default).
 * @param {string} [email]
 * @returns {TokenStorage}
 */
function getTokenStorageForUser(email) {
  const key = storeKeyForUser(email);
  if (!stores.has(key)) {
    stores.set(key, new TokenStorage({ tokenStorePath: tokenPathForUser(email) }));
  }
  return stores.get(key);
}

/**
 * Ensures the current user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh_token grant.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  const store = getTokenStorageForUser(getCurrentUser());
  const accessToken = await store.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated,
  getTokenStorageForUser
};
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/auth/ensure-authenticated.test.js`
Expected: PASS (4 tests).

- [ ] **Step 5: Run the full suite to check for regressions**

Run: `npx jest`
Expected: PASS except the known pre-existing `test/auth/token-storage.test.js › should warn...`
failure. Every other suite (including `test/email/list.test.js`,
`test/calendar/create.test.js`, which mock `../../auth`) stays green.

- [ ] **Step 6: Commit**

```bash
git add auth/index.js test/auth/ensure-authenticated.test.js
git commit -m "feat(auth): per-user TokenStorage and ensureAuthenticated"
```

---

## Task 6: `authenticate` tool issues nonce-scoped URLs

**Files:**
- Modify: `auth/tools.js`
- Test: `test/auth/tools.test.js`

**Interfaces:**
- Consumes: `getCurrentUser` from `../utils/request-context`; `createPendingLogin`
  from `./pending-logins`; `config.AUTH_CONFIG.authServerUrl`
- Produces: `handleAuthenticate(args): Promise<object>` (returns a nonce URL when a
  user context exists, legacy URL otherwise). `handleCheckAuthStatus()` unchanged
  (already user-aware via `tokenManager`). `authTools` unchanged shape.

- [ ] **Step 1: Write the failing test**

Create `test/auth/tools.test.js`:

```js
jest.mock('../../auth/pending-logins', () => ({
  createPendingLogin: jest.fn(() => 'nonce123')
}));

const { createPendingLogin } = require('../../auth/pending-logins');
const { handleAuthenticate } = require('../../auth/tools');
const { runWithUser } = require('../../utils/request-context');

describe('authenticate tool (multi-user)', () => {
  beforeEach(() => {
    delete process.env.USE_TEST_MODE;
    createPendingLogin.mockClear();
  });

  it('returns a nonce-scoped login URL for the current user', async () => {
    const res = await runWithUser('alice@x.com', () => handleAuthenticate({}));
    expect(createPendingLogin).toHaveBeenCalledWith('alice@x.com');
    expect(res.content[0].text).toContain('/auth?n=nonce123');
  });

  it('falls back to the legacy URL without a user context', async () => {
    const res = await handleAuthenticate({});
    expect(createPendingLogin).not.toHaveBeenCalled();
    expect(res.content[0].text).toContain('/auth?client_id=');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/auth/tools.test.js`
Expected: FAIL — the URL contains `/auth?client_id=` even inside a user context
(`createPendingLogin` not called).

- [ ] **Step 3: Write the implementation**

Replace the top imports + `handleAuthenticate` in `auth/tools.js`. The file becomes:

```js
/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const { getCurrentUser } = require('../utils/request-context');
const { createPendingLogin } = require('./pending-logins');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;

  // For test mode, create a test token for the current user
  if (config.USE_TEST_MODE) {
    tokenManager.createTestTokens();

    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }

  // Real authentication: build a login URL scoped to the current user.
  const email = getCurrentUser();
  let authUrl;

  if (email) {
    const nonce = createPendingLogin(email);
    authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?n=${nonce}`;
  } else {
    // No user context (stdio / single-user): legacy global flow.
    authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
  }

  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  const tokens = tokenManager.loadTokenCache();

  if (!tokens || !tokens.access_token) {
    return {
      content: [{ type: "text", text: "Not authenticated" }]
    };
  }

  return {
    content: [{ type: "text", text: "Authenticated and ready" }]
  };
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: { type: "object", properties: {}, required: [] },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: { type: "object", properties: {}, required: [] },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/auth/tools.test.js`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add auth/tools.js test/auth/tools.test.js
git commit -m "feat(auth): authenticate tool issues nonce-scoped login URLs"
```

---

## Task 7: Bind the identity header in the HTTP server

**Files:**
- Modify: `index.js`
- Test: `test/index.http.test.js`

**Interfaces:**
- Consumes: `runWithUser`, `getUserFromHeaders`, `DEFAULT_USER_HEADER` from
  `./utils/request-context`
- Produces: `createServer()`, `createHttpRequestHandler(): (req, res) => Promise<void>`,
  `startHttp()`, `startStdio()`

- [ ] **Step 1: Write the failing test**

Create `test/index.http.test.js`:

```js
process.env.USE_TEST_MODE = 'false';
process.env.OUTLOOK_TOKEN_STORE_DIR = require('fs').mkdtempSync(
  require('path').join(require('os').tmpdir(), 'mcp-http-')
);

const http = require('http');
const request = require('supertest');
const { createHttpRequestHandler } = require('../index');

describe('HTTP identity wiring', () => {
  let server;

  beforeAll(() => {
    server = http.createServer(createHttpRequestHandler());
  });

  afterAll((done) => server.close(done));

  function call(body, headers = {}) {
    return request(server)
      .post('/mcp')
      .set('Accept', 'application/json, text/event-stream')
      .set('Content-Type', 'application/json')
      .set(headers)
      .send(body);
  }

  it('serves /health', async () => {
    const res = await request(server).get('/health');
    expect(res.status).toBe(200);
    expect(res.body.status).toBe('ok');
  });

  it('routes the header user into the tool via AsyncLocalStorage', async () => {
    const res = await call(
      { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'authenticate', arguments: {} } },
      { 'x-avenu-lh-user-email': 'alice@x.com' }
    );
    expect(res.status).toBe(200);
    expect(JSON.stringify(res.body)).toContain('/auth?n=');
  });

  it('falls back to the legacy URL without the header', async () => {
    const res = await call(
      { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'authenticate', arguments: {} } }
    );
    expect(JSON.stringify(res.body)).toContain('/auth?client_id=');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/index.http.test.js`
Expected: FAIL — `createHttpRequestHandler is not a function` (and/or the server starts
listening as a side effect of `require('../index')`).

- [ ] **Step 3: Write the implementation**

Replace `index.js` with the following (transport behavior preserved; startup is now
guarded and the HTTP handler is exported for testing):

```js
#!/usr/bin/env node
/**
 * M365 Assistant MCP Server - Main entry point
 *
 * Transports:
 *   - stdio (default)  : spawned by an MCP client, e.g. Claude Desktop
 *   - http             : Streamable HTTP, set MCP_TRANSPORT=http
 *                        (used by the Docker image / docker-compose)
 */
const http = require('http');
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const config = require('./config');
const {
  DEFAULT_USER_HEADER,
  runWithUser,
  getUserFromHeaders
} = require('./utils/request-context');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools,
  ...powerAutomateTools
];

/**
 * Creates a fully configured MCP server instance.
 * Fresh per Streamable HTTP request (stateless mode).
 * @returns {Server}
 */
function createServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.fallbackRequestHandler = async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      if (method === "initialize") {
        return {
          protocolVersion: "2025-11-25",
          capabilities: { tools: {} },
          serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
        };
      }

      if (method === "tools/list") {
        return {
          tools: TOOLS.map((tool) => ({
            name: tool.name,
            description: tool.description,
            inputSchema: tool.inputSchema
          }))
        };
      }

      if (method === "resources/list") return { resources: [] };
      if (method === "prompts/list") return { prompts: [] };

      if (method === "tools/call") {
        try {
          const { name, arguments: args = {} } = params || {};
          console.error(`TOOL CALL: ${name}`);

          const tool = TOOLS.find((t) => t.name === name);
          if (tool && tool.handler) {
            return await tool.handler(args);
          }

          return {
            error: { code: -32601, message: `Tool not found: ${name}` }
          };
        } catch (error) {
          console.error(`Error in tools/call:`, error);
          return {
            error: { code: -32603, message: `Error processing tool call: ${error.message}` }
          };
        }
      }

      return { error: { code: -32601, message: `Method not found: ${method}` } };
    } catch (error) {
      console.error(`Error in fallbackRequestHandler:`, error);
      return { error: { code: -32603, message: `Error processing request: ${error.message}` } };
    }
  };

  return server;
}

async function startStdio() {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`${config.SERVER_NAME} connected and listening (stdio)`);
}

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}

/**
 * Builds the Streamable HTTP request handler (stateless).
 * @returns {(req, res) => Promise<void>}
 */
function createHttpRequestHandler() {
  const mcpPath = process.env.MCP_PATH || '/mcp';
  const allowedOrigin = process.env.MCP_ALLOWED_ORIGIN || '*';
  const userHeaderName = (process.env.MCP_USER_HEADER || DEFAULT_USER_HEADER).toLowerCase();

  return async function handleRequest(req, res) {
    res.setHeader('Access-Control-Allow-Origin', allowedOrigin);
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept, Authorization, mcp-session-id, mcp-protocol-version');
    res.setHeader('Access-Control-Expose-Headers', 'mcp-session-id');

    if (req.method === 'OPTIONS') {
      res.writeHead(204);
      res.end();
      return;
    }

    let pathname;
    try {
      pathname = new URL(req.url, `http://${req.headers.host || 'localhost'}`).pathname;
    } catch {
      res.writeHead(400, { 'Content-Type': 'text/plain' });
      res.end('Bad Request');
      return;
    }

    if (pathname === '/health' || pathname === '/healthz') {
      sendJson(res, 200, {
        status: 'ok',
        name: config.SERVER_NAME,
        version: config.SERVER_VERSION,
        tools: TOOLS.length
      });
      return;
    }

    if (pathname !== mcpPath) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not Found');
      return;
    }

    if (req.method !== 'POST') {
      res.setHeader('Allow', 'POST');
      sendJson(res, 405, {
        jsonrpc: '2.0',
        error: { code: -32000, message: 'Method not allowed. Use POST.' },
        id: null
      });
      return;
    }

    const email = getUserFromHeaders(req.headers, userHeaderName);
    const server = createServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true
    });

    res.on('close', () => {
      transport.close().catch(() => {});
      server.close().catch(() => {});
    });

    try {
      await runWithUser(email, () => transport.handleRequest(req, res));
    } catch (error) {
      console.error(`Error handling MCP request: ${error.message}`);
      if (!res.headersSent) {
        sendJson(res, 500, {
          jsonrpc: '2.0',
          error: { code: -32603, message: 'Internal server error' },
          id: null
        });
      }
    }
  };
}

async function startHttp() {
  const port = parseInt(process.env.PORT || process.env.MCP_PORT || '3000', 10);
  const host = process.env.HOST || process.env.MCP_HOST || '0.0.0.0';

  const httpServer = http.createServer(createHttpRequestHandler());
  await new Promise((resolve, reject) => {
    httpServer.once('error', reject);
    httpServer.listen(port, host, resolve);
  });

  console.error(`${config.SERVER_NAME} connected and listening (streamable http)`);
  console.error(`MCP endpoint: http://${host}:${port}${process.env.MCP_PATH || '/mcp'}`);
  console.error(`Health check: http://${host}:${port}/health`);

  const shutdown = () => {
    console.error('Shutting down HTTP server');
    httpServer.close(() => process.exit(0));
    setTimeout(() => process.exit(0), 5000).unref();
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}

function main() {
  console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
  console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

  const transportMode = (process.env.MCP_TRANSPORT || 'stdio').toLowerCase();
  if (transportMode === 'http') {
    startHttp().catch((error) => {
      console.error(`Startup error: ${error.message}`);
      process.exit(1);
    });
  } else {
    startStdio().catch((error) => {
      console.error(`Connection error: ${error.message}`);
      process.exit(1);
    });
  }
}

if (require.main === module) {
  main();
}

module.exports = { createServer, createHttpRequestHandler, startHttp, startStdio, TOOLS };
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/index.http.test.js`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add index.js test/index.http.test.js
git commit -m "feat(http): bind x-avenu-lh-user-email identity into request context"
```

---

## Task 8: Per-user consent flow in the auth server

**Files:**
- Modify: `outlook-auth-server.js`
- Test: `test/auth/auth-server.test.js`

**Interfaces:**
- Consumes: `consumePendingLogin` from `./auth/pending-logins`; `tokenPathForUser`,
  `getHomeDir` from `./auth/user-paths`
- Produces: `createAuthServer(): http.Server` (the `http.createServer(...)` instance)

- [ ] **Step 1: Write the failing test**

Create `test/auth/auth-server.test.js`:

```js
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-secret';
process.env.MS_TENANT_ID = 'test-tenant';
process.env.OUTLOOK_TOKEN_STORE_DIR = require('fs').mkdtempSync(
  require('path').join(require('os').tmpdir(), 'authsrv-')
);

const request = require('supertest');
const { createAuthServer } = require('../../outlook-auth-server');
const { createPendingLogin } = require('../../auth/pending-logins');

describe('auth server', () => {
  it('rejects an unknown nonce with 403', async () => {
    const res = await request(createAuthServer()).get('/auth?n=deadbeef');
    expect(res.status).toBe(403);
  });

  it('redirects to Microsoft for a valid nonce and consumes it', async () => {
    const nonce = createPendingLogin('alice@x.com');
    const res = await request(createAuthServer()).get(`/auth?n=${nonce}`);
    expect(res.status).toBe(302);
    expect(res.headers.location).toContain('login.microsoftonline.com');
    expect(res.headers.location).toContain('test-tenant');

    const second = await request(createAuthServer()).get(`/auth?n=${nonce}`);
    expect(second.status).toBe(403);
  });

  it('keeps the legacy single-user flow when no nonce is present', async () => {
    const res = await request(createAuthServer()).get('/auth');
    expect(res.status).toBe(302);
    expect(res.headers.location).toContain('login.microsoftonline.com');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/auth/auth-server.test.js`
Expected: FAIL — `createAuthServer is not a function` (and the module calls
`server.listen(3333)` on import).

- [ ] **Step 3: Write the implementation**

Edit `outlook-auth-server.js`:

1. Add imports near the top (after the existing `require`s):

```js
const { consumePendingLogin } = require('./auth/pending-logins');
const { tokenPathForUser } = require('./auth/user-paths');
```

2. Make the CSRF state store hold the email. Replace:

```js
// CSRF state store (state -> timestamp, cleaned up periodically)
const pendingStates = new Map();
const TEN_MINUTES = 10 * 60 * 1000;
```

with:

```js
// CSRF state store (state -> { email, createdAt }, cleaned up periodically)
const pendingStates = new Map();
const TEN_MINUTES = 10 * 60 * 1000;
```

Then update the cleanup interval body. Replace:

```js
  const now = Date.now();
  for (const [key, timestamp] of pendingStates.entries()) {
    if (now - timestamp > TEN_MINUTES) pendingStates.delete(key);
  }
```

with:

```js
  const now = Date.now();
  for (const [key, entry] of pendingStates.entries()) {
    if (now - (entry.createdAt || 0) > TEN_MINUTES) pendingStates.delete(key);
  }
```

3. In `/auth/callback`, capture the email from the state entry before deleting it.
   Replace:

```js
    pendingStates.delete(query.state);
```

with:

```js
    const stateEntry = pendingStates.get(query.state) || {};
    const stateEmail = stateEntry.email;
    pendingStates.delete(query.state);
```

(Leave the surrounding CSRF validation and error branches exactly as they are.)

4. The callback calls the exchange. Replace `exchangeCodeForTokens(query.code)`
   with `exchangeCodeForTokens(query.code, stateEmail)`.

5. Rework the `/auth` handler. Replace the body of the `else if (pathname === '/auth')`
   block up to the redirect so that a nonce is honored:

```js
  } else if (pathname === '/auth') {
    console.log('Auth request received, redirecting to Microsoft login...');

    if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
      res.writeHead(500, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head>
            <title>Configuration Error</title>
            <style>
              body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
              h1 { color: #d9534f; }
              .error-box { background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px; }
              code { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; }
            </style>
          </head>
          <body>
            <h1>Configuration Error</h1>
            <div class="error-box">
              <p>Microsoft Graph API credentials are not set. Please set the following environment variables:</p>
              <ul>
                <li><code>MS_CLIENT_ID</code></li>
                <li><code>MS_CLIENT_SECRET</code></li>
              </ul>
            </div>
          </body>
        </html>
      `);
      return;
    }

    // Resolve the user this consent belongs to. With a nonce, it is the
    // authenticated user; without one, this is the legacy single-user flow.
    const userEmail = parsedUrl.query.n ? consumePendingLogin(parsedUrl.query.n) : null;
    if (parsedUrl.query.n && !userEmail) {
      console.error('Invalid or expired login nonce');
      res.writeHead(403, { 'Content-Type': 'text/html' });
      res.end('<html><body><h1>Invalid or expired login link</h1><p>Please start authentication again.</p></body></html>');
      return;
    }

    // Generate cryptographically secure state parameter for CSRF protection
    const state = crypto.randomBytes(32).toString('hex');
    pendingStates.set(state, { email: userEmail, createdAt: Date.now() });

    const authParams = {
      client_id: AUTH_CONFIG.clientId,
      response_type: 'code',
      redirect_uri: AUTH_CONFIG.redirectUri,
      scope: AUTH_CONFIG.scopes.join(' '),
      response_mode: 'query',
      state
    };

    const authUrl = `${AUTH_CONFIG.authorityHost}/${AUTH_CONFIG.tenantId}/oauth2/v2.0/authorize?${querystring.stringify(authParams)}`;
    console.log(`Redirecting to: ${authUrl}`);

    res.writeHead(302, { 'Location': authUrl });
    res.end();
  } else if (pathname === '/') {
```

6. Change `exchangeCodeForTokens` to accept an email and write to that user's file.
   Replace the function signature and the write block:

```js
function exchangeCodeForTokens(code, email) {
  return new Promise((resolve, reject) => {
    const postData = querystring.stringify({
      client_id: AUTH_CONFIG.clientId,
      client_secret: AUTH_CONFIG.clientSecret,
      code: code,
      redirect_uri: AUTH_CONFIG.redirectUri,
      grant_type: 'authorization_code',
      scope: AUTH_CONFIG.scopes.join(' ')
    });

    const options = {
      hostname: AUTH_CONFIG.authorityHost.replace(/^https?:\/\//, '').split('/')[0],
      path: `/${AUTH_CONFIG.tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse = JSON.parse(data);
            tokenResponse.expires_at = Date.now() + (tokenResponse.expires_in * 1000);

            const tokenPath = tokenPathForUser(email);
            fs.mkdirSync(path.dirname(tokenPath), { recursive: true });
            fs.writeFileSync(tokenPath, JSON.stringify(tokenResponse, null, 2), { encoding: 'utf8', mode: 0o600 });
            console.log(`Tokens saved to ${tokenPath}`);

            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error parsing token response: ${error.message}`));
          }
        } else {
          reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });

    req.on('error', (error) => {
      reject(error);
    });

    req.write(postData);
    req.end();
  });
}
```

7. Extract server creation and guard the listener.
   - **7a.** Insert a line `function createAuthServer() {` immediately **before** the
     `const server = http.createServer((req, res) => {` line (right after the
     `// Create HTTP server` comment). Do not re-indent the existing body.
   - **7b.** Immediately after the closing `});` that ends the
     `http.createServer((req, res) => { ... })` call (the `});` just before
     `function exchangeCodeForTokens`), insert:

```js
  return server;
}
```

   - **7c.** Replace the entire `// Start server` block — the `server.listen(...)`
     call and both `process.on('SIGINT'/'SIGTERM', ...)` handlers. It currently reads:

```js
// Start server
server.listen(AUTH_PORT, () => {
  console.log(`Authentication server running at http://localhost:${AUTH_PORT}`);
  console.log(`Waiting for authentication callback at ${AUTH_CONFIG.redirectUri}`);
  console.log(`Token will be stored at: ${tokenPathForUser(null)}`);

  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.log('\n⚠️  WARNING: Microsoft Graph API credentials are not set.');
    console.log('   Please set the MS_CLIENT_ID and MS_CLIENT_SECRET environment variables.');
  }
});

// Handle termination
process.on('SIGINT', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});
```

with:

```js
function startAuthServer() {
  const server = createAuthServer();
  server.listen(AUTH_PORT, () => {
    console.log(`Authentication server running at http://localhost:${AUTH_PORT}`);
    console.log(`Waiting for authentication callback at ${AUTH_CONFIG.redirectUri}`);
    console.log(`Token will be stored at: ${tokenPathForUser(null)}`);

    if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
      console.log('\n⚠️  WARNING: Microsoft Graph API credentials are not set.');
      console.log('   Please set the MS_CLIENT_ID and MS_CLIENT_SECRET environment variables.');
    }
  });
  return server;
}

if (require.main === module) {
  startAuthServer();
}

// Handle termination
process.on('SIGINT', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});

module.exports = { createAuthServer };
```

Make sure `path` is required at the top (it already is: `const path = require('path');`).
Make sure `tokenPathForUser` is imported (step 1).

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/auth/auth-server.test.js`
Expected: PASS (3 tests).

- [ ] **Step 5: Sanity-check the script still starts**

Run: `AUTH_PORT=41999 MS_CLIENT_ID=x MS_CLIENT_SECRET=y MS_TENANT_ID=t node outlook-auth-server.js & sleep 1; curl -s -o /dev/null -w "%{http_code}\n" http://127.0.0.1:41999/; kill %1`
Expected: prints `200`.

- [ ] **Step 6: Commit**

```bash
git add outlook-auth-server.js test/auth/auth-server.test.js
git commit -m "feat(auth-server): nonce-scoped, per-user consent flow"
```

---

## Task 9: Docker, env, and docs

**Files:**
- Modify: `docker-compose.yml`
- Modify: `.env.example`
- Modify: `README.md`

**Interfaces:** none (configuration/docs).

- [ ] **Step 1: Point both services at the store root**

In `docker-compose.yml`, replace, in **both** the `mcp` and `auth` services, the line:

```yaml
      OUTLOOK_TOKEN_STORE_PATH: /data/.outlook-mcp-tokens.json
```

with:

```yaml
      OUTLOOK_TOKEN_STORE_DIR: /data
```

(The `HOME: /data` lines stay.)

- [ ] **Step 2: Document env vars in `.env.example`**

Append to `.env.example`:

```bash
# --- Multi-user (optional) ---
# Header (trusted) that carries the user's identity, per request.
# MCP_USER_HEADER=x-avenu-lh-user-email
# Root directory for per-user tokens (tokens/) and pending logins (pending/).
# OUTLOOK_TOKEN_STORE_DIR=/data
```

- [ ] **Step 3: Add a "Multi-user" section to the README**

Add under the Docker section (after the Docker configuration table):

```markdown
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
```

- [ ] **Step 4: Verify compose is valid**

Run: `docker compose config >/dev/null && echo COMPOSE_OK`
Expected: prints `COMPOSE_OK`.

- [ ] **Step 5: Commit**

```bash
git add docker-compose.yml .env.example README.md
git commit -m "docs(docker): configure and document multi-user token store"
```

---

## Task 10: End-to-end manual verification

**Files:** none (manual verification only).

- [ ] **Step 1: Full test suite**

Run: `npx jest`
Expected: everything green **except** the known pre-existing
`test/auth/token-storage.test.js › should warn...` failure.

- [ ] **Step 2: Build and run the image**

```bash
docker build -t outlook-mcp:latest .
docker rm -f omcp >/dev/null 2>&1
docker run -d --name omcp -p 41234:3000 \
  -e USE_TEST_MODE=false \
  -e MS_CLIENT_ID=dummy -e MS_CLIENT_SECRET=dummy -e MS_TENANT_ID=dummy \
  outlook-mcp:latest
sleep 2
```

- [ ] **Step 3: Verify per-user identity reaches the tools**

```bash
curl -s http://127.0.0.1:41234/health; echo

# With a user header: authenticate returns a nonce-scoped URL
curl -s -X POST http://127.0.0.1:41234/mcp \
  -H 'Content-Type: application/json' \
  -H 'Accept: application/json, text/event-stream' \
  -H 'x-avenu-lh-user-email: alice@x.com' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"authenticate","arguments":{}}}'
echo

# Without the header: legacy URL
curl -s -X POST http://127.0.0.1:41234/mcp \
  -H 'Content-Type: application/json' \
  -H 'Accept: application/json, text/event-stream' \
  -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"authenticate","arguments":{}}}'
echo
```

Expected: first response contains `/auth?n=`; second contains `/auth?client_id=`.
Confirm the nonce file exists:

```bash
docker exec omcp ls /data/pending
```

- [ ] **Step 4: Verify per-user isolation in the container**

```bash
docker exec omcp node -e "
const rc=require('/app/utils/request-context');
const tm=require('/app/auth/token-manager');
const f=Date.now()+3600000;
rc.runWithUser('alice@x.com',()=>tm.saveTokenCache({access_token:'A',expires_at:f}));
rc.runWithUser('bob@x.com',()=>tm.saveTokenCache({access_token:'B',expires_at:f}));
console.log('alice:', rc.runWithUser('alice@x.com',()=>tm.getAccessToken()));
console.log('bob:', rc.runWithUser('bob@x.com',()=>tm.getAccessToken()));
"
docker exec omcp ls /data/tokens
```

Expected: `alice: A`, `bob: B`, and two files under `/data/tokens`.

- [ ] **Step 5: Clean up**

```bash
docker rm -f omcp
```

- [ ] **Step 6: No commit** (verification only).

---

## Self-Review Notes

- **Spec coverage:** identity header (Task 7), ALS propagation (Tasks 1, 7), per-user
  storage layout (Task 2), token manager (Task 4), per-user TokenStorage/refresh
  (Task 5), nonce consent (Tasks 3, 6, 8), config/Docker/backward-compat (Tasks 5, 9),
  testing (each task + Task 10). Non-goals require no tasks.
- **Type consistency:** `tokenPathForUser`, `currentTokenPath`, `userKey`,
  `pendingLoginPath`, `createPendingLogin`, `consumePendingLogin`,
  `getTokenStorageForUser`, `getUserFromHeaders`, `runWithUser`, `getCurrentUser`,
  `createAuthServer`, `createHttpRequestHandler` are used with the same names and
  shapes in every task that references them.
- **Pre-existing failure** is called out globally and must not be "fixed" here.
