jest.mock('../../auth/token-storage');

const TokenStorage = require('../../auth/token-storage');
const { ensureAuthenticated, getTokenStorageForUser } = require('../../auth');
const { runWithUser } = require('../../utils/request-context');

describe('ensureAuthenticated (multi-user)', () => {
  it('throws "Authentication required" when no token is available', async () => {
    TokenStorage.mockImplementation(() => ({
      getValidAccessToken: jest.fn().mockResolvedValue(null)
    }));
    await expect(runWithUser('nobody@x.com', () => ensureAuthenticated()))
      .rejects.toThrow('Authentication required');
  });

  it('returns the current user\'s access token', async () => {
    TokenStorage.mockImplementation(() => ({
      getValidAccessToken: jest.fn().mockResolvedValue('token-123')
    }));
    const token = await runWithUser('alice@x.com', () => ensureAuthenticated());
    expect(token).toBe('token-123');
  });

  it('throws when forceNew is requested', async () => {
    await expect(ensureAuthenticated(true)).rejects.toThrow('Authentication required');
  });

  it('caches exactly one store per user', () => {
    const a1 = getTokenStorageForUser('alice@x.com');
    const a2 = getTokenStorageForUser('alice@x.com');
    const b1 = getTokenStorageForUser('bob@x.com');
    expect(a1).toBe(a2);
    expect(a1).not.toBe(b1);
  });
});
