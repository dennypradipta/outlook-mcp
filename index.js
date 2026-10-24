#!/usr/bin/env node
/**
 * M365 Assistant MCP Server - Main entry point
 *
 * Transports:
 *   - stdio (default)  : spawned by an MCP client, e.g. Claude Desktop
 *   - http             : Streamable HTTP, set MCP_TRANSPORT=http
 *                        (used by the Docker image / docker-compose)
 */
const http = require('http');
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const config = require('./config');
const {
  DEFAULT_USER_HEADER,
  runWithUser,
  getUserFromHeaders
} = require('./utils/request-context');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools,
  ...powerAutomateTools
];

/**
 * Creates a fully configured MCP server instance.
 * Fresh per Streamable HTTP request (stateless mode).
 * @returns {Server}
 */
function createServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.fallbackRequestHandler = async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      if (method === "initialize") {
        return {
          protocolVersion: "2025-11-25",
          capabilities: { tools: {} },
          serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
        };
      }

      if (method === "tools/list") {
        return {
          tools: TOOLS.map((tool) => ({
            name: tool.name,
            description: tool.description,
            inputSchema: tool.inputSchema
          }))
        };
      }

      if (method === "resources/list") return { resources: [] };
      if (method === "prompts/list") return { prompts: [] };

      if (method === "tools/call") {
        try {
          const { name, arguments: args = {} } = params || {};
          console.error(`TOOL CALL: ${name}`);

          const tool = TOOLS.find((t) => t.name === name);
          if (tool && tool.handler) {
            return await tool.handler(args);
          }

          return {
            error: { code: -32601, message: `Tool not found: ${name}` }
          };
        } catch (error) {
          console.error(`Error in tools/call:`, error);
          return {
            error: { code: -32603, message: `Error processing tool call: ${error.message}` }
          };
        }
      }

      return { error: { code: -32601, message: `Method not found: ${method}` } };
    } catch (error) {
      console.error(`Error in fallbackRequestHandler:`, error);
      return { error: { code: -32603, message: `Error processing request: ${error.message}` } };
    }
  };

  return server;
}

async function startStdio() {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`${config.SERVER_NAME} connected and listening (stdio)`);
}

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}

/**
 * Builds the Streamable HTTP request handler (stateless).
 * @returns {(req, res) => Promise<void>}
 */
function createHttpRequestHandler() {
  const mcpPath = process.env.MCP_PATH || '/mcp';
  const allowedOrigin = process.env.MCP_ALLOWED_ORIGIN || '*';
  const userHeaderName = (process.env.MCP_USER_HEADER || DEFAULT_USER_HEADER).toLowerCase();

  return async function handleRequest(req, res) {
    res.setHeader('Access-Control-Allow-Origin', allowedOrigin);
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept, Authorization, mcp-session-id, mcp-protocol-version');
    res.setHeader('Access-Control-Expose-Headers', 'mcp-session-id');

    if (req.method === 'OPTIONS') {
      res.writeHead(204);
      res.end();
      return;
    }

    let pathname;
    try {
      pathname = new URL(req.url, `http://${req.headers.host || 'localhost'}`).pathname;
    } catch {
      res.writeHead(400, { 'Content-Type': 'text/plain' });
      res.end('Bad Request');
      return;
    }

    if (pathname === '/health' || pathname === '/healthz') {
      sendJson(res, 200, {
        status: 'ok',
        name: config.SERVER_NAME,
        version: config.SERVER_VERSION,
        tools: TOOLS.length
      });
      return;
    }

    if (pathname !== mcpPath) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not Found');
      return;
    }

    if (req.method !== 'POST') {
      res.setHeader('Allow', 'POST');
      sendJson(res, 405, {
        jsonrpc: '2.0',
        error: { code: -32000, message: 'Method not allowed. Use POST.' },
        id: null
      });
      return;
    }

    const email = getUserFromHeaders(req.headers, userHeaderName);
    const server = createServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true
    });

    res.on('close', () => {
      transport.close().catch(() => {});
      server.close().catch(() => {});
    });

    try {
      await server.connect(transport);
      await runWithUser(email, () => transport.handleRequest(req, res));
    } catch (error) {
      console.error(`Error handling MCP request: ${error.message}`);
      if (!res.headersSent) {
        sendJson(res, 500, {
          jsonrpc: '2.0',
          error: { code: -32603, message: 'Internal server error' },
          id: null
        });
      }
    }
  };
}

async function startHttp() {
  const port = parseInt(process.env.PORT || process.env.MCP_PORT || '3000', 10);
  const host = process.env.HOST || process.env.MCP_HOST || '0.0.0.0';

  const httpServer = http.createServer(createHttpRequestHandler());
  await new Promise((resolve, reject) => {
    httpServer.once('error', reject);
    httpServer.listen(port, host, resolve);
  });

  console.error(`${config.SERVER_NAME} connected and listening (streamable http)`);
  console.error(`MCP endpoint: http://${host}:${port}${process.env.MCP_PATH || '/mcp'}`);
  console.error(`Health check: http://${host}:${port}/health`);

  const shutdown = () => {
    console.error('Shutting down HTTP server');
    httpServer.close(() => process.exit(0));
    setTimeout(() => process.exit(0), 5000).unref();
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}

function main() {
  console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
  console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

  const transportMode = (process.env.MCP_TRANSPORT || 'stdio').toLowerCase();
  if (transportMode === 'http') {
    startHttp().catch((error) => {
      console.error(`Startup error: ${error.message}`);
      process.exit(1);
    });
  } else {
    startStdio().catch((error) => {
      console.error(`Connection error: ${error.message}`);
      process.exit(1);
    });
  }
}

if (require.main === module) {
  main();
}

module.exports = { createServer, createHttpRequestHandler, startHttp, startStdio, TOOLS };
