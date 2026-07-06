# Build stage
FROM node:24-alpine AS builder

RUN corepack enable && corepack prepare pnpm@11.5.1 --activate

WORKDIR /app

# Workspace manifests first for a cached install layer
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
COPY packages/core/package.json ./packages/core/
COPY packages/graph/package.json ./packages/graph/
RUN pnpm install --frozen-lockfile

# Build core then the server (-r = topological), so the server bundles core's dist.
# Then produce a self-contained production deployment (prod-only node_modules + the
# package's published files) outside the workspace.
COPY . .
RUN pnpm -r build
RUN pnpm --filter microsoft-mcp-server deploy --prod /prod

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
