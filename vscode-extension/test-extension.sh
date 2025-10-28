#!/bin/bash
# Quick test script for the ExcelMcp VS Code extension

set -e

echo "=================================="
echo "ExcelMcp Extension Quick Test"
echo "=================================="
echo ""

# Check if we're in the right directory
if [ ! -f "package.json" ]; then
    echo "❌ Error: Not in vscode-extension directory"
    echo "   Run: cd vscode-extension"
    exit 1
fi

echo "✅ In correct directory"
echo ""

# Check Node.js
echo "Checking Node.js..."
if ! command -v node &> /dev/null; then
    echo "❌ Node.js not found. Please install Node.js 20+"
    exit 1
fi
NODE_VERSION=$(node --version)
echo "✅ Node.js: $NODE_VERSION"
echo ""

# Check npm
echo "Checking npm..."
if ! command -v npm &> /dev/null; then
    echo "❌ npm not found"
    exit 1
fi
NPM_VERSION=$(npm --version)
echo "✅ npm: $NPM_VERSION"
echo ""

# Install dependencies
echo "Installing dependencies..."
if [ ! -d "node_modules" ]; then
    npm install
else
    echo "✅ Dependencies already installed"
fi
echo ""

# Compile TypeScript
echo "Compiling TypeScript..."
npm run compile
if [ $? -eq 0 ]; then
    echo "✅ TypeScript compilation successful"
else
    echo "❌ TypeScript compilation failed"
    exit 1
fi
echo ""

# Check compiled output
if [ -f "out/extension.js" ]; then
    SIZE=$(ls -lh out/extension.js | awk '{print $5}')
    echo "✅ Compiled output: out/extension.js ($SIZE)"
else
    echo "❌ Compiled output not found"
    exit 1
fi
echo ""

# Run linter
echo "Running ESLint..."
npm run lint
if [ $? -eq 0 ]; then
    echo "✅ ESLint passed"
else
    echo "⚠️  ESLint warnings (check output above)"
fi
echo ""

# Package extension
echo "Packaging extension..."
echo "y" | npx @vscode/vsce package --no-dependencies --allow-missing-repository > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "✅ Extension packaged successfully"
else
    echo "❌ Packaging failed"
    exit 1
fi
echo ""

# Check VSIX
if [ -f "excelmcp-1.0.0.vsix" ]; then
    SIZE=$(ls -lh excelmcp-1.0.0.vsix | awk '{print $5}')
    echo "✅ VSIX created: excelmcp-1.0.0.vsix ($SIZE)"
else
    echo "❌ VSIX file not found"
    exit 1
fi
echo ""

# List VSIX contents
echo "VSIX contents:"
unzip -l excelmcp-1.0.0.vsix | grep -E "extension/(package.json|out/extension.js|icon.png|readme.md)"
echo ""

echo "=================================="
echo "✅ All tests passed!"
echo "=================================="
echo ""
echo "Next steps:"
echo "1. Install in VS Code: Ctrl+Shift+P → 'Install from VSIX'"
echo "2. Select: excelmcp-1.0.0.vsix"
echo "3. Verify: Ask GitHub Copilot to list Excel MCP tools"
echo ""
