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
