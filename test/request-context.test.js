const {
  runWithUser,
  getCurrentUser,
  getUserFromHeaders,
  DEFAULT_USER_HEADER
} = require('../utils/request-context');

describe('request-context', () => {
  it('returns undefined outside a bound context', () => {
    expect(getCurrentUser()).toBeUndefined();
  });

  it('returns the bound user inside runWithUser', () => {
    const result = runWithUser('alice@x.com', () => getCurrentUser());
    expect(result).toBe('alice@x.com');
  });

  it('isolates interleaved async contexts', async () => {
    const results = {};
    await Promise.all([
      runWithUser('a@x.com', async () => {
        await new Promise((r) => setTimeout(r, 20));
        results.a = getCurrentUser();
      }),
      runWithUser('b@x.com', async () => {
        await new Promise((r) => setTimeout(r, 5));
        results.b = getCurrentUser();
      })
    ]);
    expect(results.a).toBe('a@x.com');
    expect(results.b).toBe('b@x.com');
  });

  it('extracts and trims a header value (case-insensitive key)', () => {
    expect(getUserFromHeaders({ 'x-avenu-lh-user-email': ' alice@x.com ' })).toBe('alice@x.com');
  });

  it('returns undefined for missing or blank header', () => {
    expect(getUserFromHeaders({})).toBeUndefined();
    expect(getUserFromHeaders({ 'x-avenu-lh-user-email': '   ' })).toBeUndefined();
  });

  it('handles array header values', () => {
    expect(getUserFromHeaders({ 'x-avenu-lh-user-email': ['a@x.com', 'b@x.com'] })).toBe('a@x.com');
  });

  it('honors a custom header name', () => {
    expect(getUserFromHeaders({ 'x-user': 'c@x.com' }, 'x-user')).toBe('c@x.com');
    expect(DEFAULT_USER_HEADER).toBe('x-avenu-lh-user-email');
  });
});
