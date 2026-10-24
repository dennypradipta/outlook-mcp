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
