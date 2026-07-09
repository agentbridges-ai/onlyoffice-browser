#!/bin/bash

# Exit on error
set -e

echo "Starting build process..."

# Run Vite build through the project-local binary.
./node_modules/.bin/vite build

# Keep the deployable demo/runtime lean by removing low-frequency bundled
# assets and emitting canonical path packs for per-document-type CDN sync.
node scripts/build-onlyoffice-runtime-assets.mjs \
    --input dist \
    --prune-root \
    --split-output .onlyoffice-runtime-asset-packs

node scripts/patch-onlyoffice-print-fallback.mjs dist

# Inject timestamp into sw.js for versioning
SW_PATH="dist/sw.js"
if [ -f "$SW_PATH" ]; then
    TIMESTAMP=$(date +%s)
    # Use sed to replace the placeholder with the actual timestamp
    # Handling cross-platform sed (macOS vs Linux)
    if [[ "$OSTYPE" == "darwin"* ]]; then
        sed -i '' "s/SW_VERSION_PLACEHOLDER/$TIMESTAMP/g" "$SW_PATH"
    else
        sed -i "s/SW_VERSION_PLACEHOLDER/$TIMESTAMP/g" "$SW_PATH"
    fi
    echo "Service Worker version updated with timestamp: $TIMESTAMP"
else
    echo "Warning: dist/sw.js not found, skipping version injection."
fi

echo "Build completed successfully!"
