const fs = require('fs');
const os = require('os');
const path = require('path');

describe('token-manager (multi-user)', () => {
  const originalEnv = process.env;
  let tmp;
  let tokenManager;
  let requestContext;
  let userPaths;

  beforeEach(() => {
    jest.resetModules();
    process.env = { ...originalEnv };
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'tokmgr-'));
    process.env.OUTLOOK_TOKEN_STORE_DIR = tmp;
    tokenManager = require('../../auth/token-manager');
    requestContext = require('../../utils/request-context');
    userPaths = require('../../auth/user-paths');
  });

  afterAll(() => {
    process.env = originalEnv;
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
    process.env.HOME = tmp;
    delete process.env.OUTLOOK_TOKEN_STORE_DIR;
    tokenManager.saveTokenCache({ access_token: 'global-token', expires_at: future });
    expect(fs.existsSync(path.join(tmp, '.outlook-mcp-tokens.json'))).toBe(true);
  });
});
