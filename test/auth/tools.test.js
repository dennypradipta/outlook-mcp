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
