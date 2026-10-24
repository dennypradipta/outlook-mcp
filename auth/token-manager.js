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
