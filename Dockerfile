# syntax=docker/dockerfile:1

FROM node:22-alpine AS base
ENV NODE_ENV=production
WORKDIR /app

# Install dependencies first for better layer caching.
# Uses the lockfile with `npm ci` for reproducible builds.
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --no-audit --no-fund && npm cache clean --force

# Copy application source
COPY . .

# Run as the unprivileged built-in `node` user.
# HOME=/data so OAuth tokens are written to the mounted volume.
ENV HOME=/data \
    MCP_TRANSPORT=http \
    HOST=0.0.0.0 \
    PORT=3000 \
    MCP_PATH=/mcp \
    OUTLOOK_TOKEN_STORE_DIR=/data

RUN mkdir -p /data && chown -R node:node /data /app
USER node

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:'+(process.env.PORT||3000)+'/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

# Default: run the MCP server over Streamable HTTP.
# Override the command to run the OAuth callback server instead:
#   docker run ... <image> node outlook-auth-server.js
CMD ["node", "index.js"]
