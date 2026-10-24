const fs = require('fs');
const os = require('os');
const path = require('path');

describe('pending-logins', () => {
  const originalEnv = process.env;
  let tmp;
  let pendingLogins;
  let userPaths;

  beforeEach(() => {
    jest.resetModules();
    process.env = { ...originalEnv };
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pending-'));
    process.env.OUTLOOK_TOKEN_STORE_DIR = tmp;
    pendingLogins = require('../../auth/pending-logins');
    userPaths = require('../../auth/user-paths');
  });

  afterAll(() => {
    process.env = originalEnv;
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
