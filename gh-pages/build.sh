#!/bin/bash
# Jekyll build script for Excel MCP Server documentation
# This script copies shared content files before building Jekyll
# Used by both local development and GitHub Actions

set -e  # Exit on error

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"

echo "📁 Copying shared content files..."

# Create _includes directory if it doesn't exist
mkdir -p "$SCRIPT_DIR/_includes"

# Copy FEATURES.md from root
cp "$ROOT_DIR/FEATURES.md" "$SCRIPT_DIR/_includes/features.md"
echo "   ✓ Copied FEATURES.md"

# Copy CHANGELOG.md from vscode-extension
cp "$ROOT_DIR/vscode-extension/CHANGELOG.md" "$SCRIPT_DIR/_includes/changelog.md"
echo "   ✓ Copied CHANGELOG.md"

# Copy INSTALLATION.md from docs
cp "$ROOT_DIR/docs/INSTALLATION.md" "$SCRIPT_DIR/_includes/installation.md"
echo "   ✓ Copied INSTALLATION.md"

# Determine build mode
if [ "$1" == "serve" ]; then
    echo ""
    echo "🚀 Starting Jekyll server..."
    cd "$SCRIPT_DIR"
    bundle exec jekyll serve --host 127.0.0.1 --port 4000
elif [ "$1" == "production" ] || [ "$JEKYLL_ENV" == "production" ]; then
    echo ""
    echo "🏗️  Building for production..."
    cd "$SCRIPT_DIR"
    JEKYLL_ENV=production bundle exec jekyll build
else
    echo ""
    echo "🏗️  Building for development..."
    cd "$SCRIPT_DIR"
    bundle exec jekyll build
fi

echo ""
echo "✅ Build complete!"
