#!/bin/bash
set -euo pipefail

VERSION=$(node -p "require('./package.json').version")
MANIFEST="${1:-manifest.json}"
SUFFIX="${2:-}"

if [ -n "$SUFFIX" ]; then
  BUNDLE_NAME="microsoft365-mcp-server-${VERSION}-${SUFFIX}.mcpb"
else
  BUNDLE_NAME="microsoft365-mcp-server-${VERSION}.mcpb"
fi

if [ ! -f "$MANIFEST" ]; then
  echo "Error: $MANIFEST not found"
  exit 1
fi

rm -rf .mcpb-build
mkdir -p .mcpb-build

cp "$MANIFEST" .mcpb-build/manifest.json

cd .mcpb-build
zip -9 "../${BUNDLE_NAME}" manifest.json
cd ..

rm -rf .mcpb-build

echo "Built: ${BUNDLE_NAME}"
