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
# Strip top block (H1 title, bold subtitle, hr/blank) and convert remaining H1 to H2
awk '
    BEGIN { inheader=0; headerdone=0 }
    {
        if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
        if (inheader==1 && /^\*\*/) { next }                            # drop bold subtitle line
        if (inheader==1 && /^---/) { inheader=0; headerdone=1; next }    # drop hr then end header
        if (inheader==1 && /^$/) { next }                                # skip blank lines while in header
        if (inheader==1) { next }                                        # drop any lingering header lines
        if (/^$/ && headerdone==0) { next }                              # drop leading blanks before content
        if (/^# /) { sub(/^# /, "## "); print; next }                   # convert any remaining H1 → H2
        print
    }
' "$ROOT_DIR/FEATURES.md" > "$SCRIPT_DIR/_includes/features.md"
echo "   ✓ Copied FEATURES.md (stripped top block, H1→H2)"

# Copy CHANGELOG.md from vscode-extension
# Strip top H1 block (title + paragraph) and convert remaining H1 to H2
awk '
    BEGIN { inheader=0; headerdone=0 }
    {
        if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
        if (inheader==1 && /^All notable/) { next }                      # drop description line
        if (inheader==1 && /^$/) { inheader=0; headerdone=1; next }      # blank line ends header
        if (/^# /) { sub(/^# /, "## "); print; next }                   # convert any remaining H1 → H2
        print
    }
' "$ROOT_DIR/vscode-extension/CHANGELOG.md" > "$SCRIPT_DIR/_includes/changelog.md"
echo "   ✓ Copied CHANGELOG.md (stripped top H1 block, H1→H2)"

# Copy INSTALLATION.md from docs
# Strip top H1 block (title + paragraph)
awk '
    BEGIN { inheader=0; headerdone=0 }
    {
        if (headerdone==0 && /^# /) { inheader=1; next }                 # drop H1 title
        if (inheader==1 && /^Complete installation/) { next }            # drop description line
        if (inheader==1 && /^$/) { inheader=0; headerdone=1; next }      # blank line ends header
        print
    }
' "$ROOT_DIR/docs/INSTALLATION.md" > "$SCRIPT_DIR/_includes/installation.md"
echo "   ✓ Copied INSTALLATION.md (stripped top H1 block)"

# Copy CONTRIBUTING.md from docs
cp "$ROOT_DIR/docs/CONTRIBUTING.md" "$SCRIPT_DIR/_includes/contributing.md"
echo "   ✓ Copied CONTRIBUTING.md"

# Copy SECURITY.md from docs
cp "$ROOT_DIR/docs/SECURITY.md" "$SCRIPT_DIR/_includes/security.md"
echo "   ✓ Copied SECURITY.md"

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
