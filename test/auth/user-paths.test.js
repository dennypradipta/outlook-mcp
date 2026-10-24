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
