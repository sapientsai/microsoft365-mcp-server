# Build stage
FROM node:24-alpine AS builder

RUN corepack enable && corepack prepare pnpm@11.5.1 --activate

WORKDIR /app

# Workspace manifests first for a cached install layer
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
COPY packages/microsoft365/package.json ./packages/microsoft365/
RUN pnpm install --frozen-lockfile

# Build the server, then produce a self-contained production deployment
# (prod-only node_modules + the package's published files) outside the workspace.
COPY . .
RUN pnpm --filter microsoft365-mcp-server build
RUN pnpm --filter microsoft365-mcp-server deploy --prod /prod

# Production stage
FROM node:24-alpine AS production

ARG GIT_HASH=""
ENV GIT_HASH=${GIT_HASH}

WORKDIR /app

COPY --from=builder /prod ./

ENV NODE_ENV=production
ENV PORT=8080
ENV TRANSPORT_TYPE=httpStream
ENV FASTMCP_HOST=0.0.0.0

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
  CMD wget --no-verbose --tries=1 --spider http://127.0.0.1:${PORT}/ping || exit 1

CMD ["node", "dist/bin.js"]
