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
