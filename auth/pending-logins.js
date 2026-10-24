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
