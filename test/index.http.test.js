process.env.USE_TEST_MODE = 'false';
process.env.OUTLOOK_TOKEN_STORE_DIR = require('fs').mkdtempSync(
  require('path').join(require('os').tmpdir(), 'mcp-http-')
);

const http = require('http');
const request = require('supertest');
const { createHttpRequestHandler } = require('../index');

// supertest starts/closes this server as needed for each request.
const server = http.createServer(createHttpRequestHandler());

describe('HTTP identity wiring', () => {
  function call(body, headers = {}) {
    return request(server)
      .post('/mcp')
      .set('Accept', 'application/json, text/event-stream')
      .set('Content-Type', 'application/json')
      .set(headers)
      .send(body);
  }

  it('serves /health', async () => {
    const res = await request(server).get('/health');
    expect(res.status).toBe(200);
    expect(res.body.status).toBe('ok');
  });

  it('routes the header user into the tool via AsyncLocalStorage', async () => {
    const res = await call(
      { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'authenticate', arguments: {} } },
      { 'x-avenu-lh-user-email': 'alice@x.com' }
    );
    expect(res.status).toBe(200);
    expect(JSON.stringify(res.body)).toContain('/auth?n=');
  });

  it('falls back to the legacy URL without the header', async () => {
    const res = await call(
      { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'authenticate', arguments: {} } }
    );
    expect(JSON.stringify(res.body)).toContain('/auth?client_id=');
  });
});
