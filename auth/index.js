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
