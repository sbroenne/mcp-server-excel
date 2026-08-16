---
applyTo: "vscode-extension/**"
---

# VS Code Extension Development Instructions

> **Instructions for developing the ExcelMcp VS Code Extension**

## Extension Overview

The ExcelMcp VS Code Extension provides one-click installation of the ExcelMcp MCP server for Visual Studio Code, enabling AI assistants like GitHub Copilot to automate Microsoft Excel.

**Key Files:**
- `package.json` - Extension manifest (metadata, dependencies, version)
- `src/extension.ts` - Extension entry point (activation, MCP registration)
- `README.md` - Marketplace description page
- `CHANGELOG.md` - Version history (marketplace changelog tab)
- `DEVELOPMENT.md` - Developer guide
- `icon.png` - Extension icon (128x128px, displayed in marketplace)

---

## Changelog and Changesets (CRITICAL)

Never edit the root `CHANGELOG.md` manually. It is generated from changeset
fragments during release.

For a user-visible extension change:

1. Run `npx changeset`.
2. Select the appropriate bump type.
3. Write an end-user-facing summary.
4. Commit the generated `.changeset/<name>.md` with the pull request.

The unified release workflow compiles pending changesets into `CHANGELOG.md`,
updates package versions, builds every deliverable, and publishes the release.

---

## Version Management

### Automatic Version Management (Unified Release Workflow)

**DO NOT manually edit package.json version** - The unified release workflow handles this:

Run **Release All Components** with a semantic version bump or custom version:

```powershell
gh workflow run release.yml -f version_bump=patch
```

Unified workflow automatically:
1. Calculates the version from the latest tag and workflow input
2. Updates `package.json` version using `npm version`
3. Compiles changesets into changelog and release notes
4. Builds and packages VS Code extension
5. Builds all other components (MCP Server, CLI, MCPB)
6. Publishes to VS Code Marketplace and NuGet
7. Creates unified GitHub release with all artifacts

### Local Testing (Manual Version Bump)

For local testing only, use npm version commands:

```powershell
npm version patch   # 1.0.0 → 1.0.1
npm version minor   # 1.0.0 → 1.1.0
npm version major   # 1.0.0 → 2.0.0
```

**Important:** Don't commit manual version changes - they're for testing only.

---

## Marketplace Information

### What Users See

**VS Code Marketplace displays:**

1. **package.json metadata**:
   - `displayName` - Title shown in marketplace
   - `description` - Subtitle/summary
   - `icon` - Extension icon (128x128px minimum)
   - `categories` - Marketplace categories
   - `keywords` - Search terms
   - `publisher` - Publisher ID

2. **README.md** - Main description page (features, installation, docs)
3. **CHANGELOG.md** - Changelog tab in marketplace
4. **LICENSE** - License information

### Critical Files for Marketplace

- ✅ **README.md** - Keep up-to-date with accurate commands and features
- ✅ **CHANGELOG.md** - Maintain version history
- ✅ **package.json** - Ensure metadata is accurate
- ✅ **icon.png** - High-quality 128x128px PNG

---

## Bundled MCP Server

The extension packages a self-contained Windows MCP Server under `bin/`.
`src/extension.ts` registers that bundled executable through VS Code's MCP
definition provider, so users do not need a .NET runtime or separate tool install.

When changing packaging, verify `package.json` scripts, `.vscodeignore`,
`src/extension.ts`, the VSIX contents, and `README.md` stay aligned.

---

## Development Workflow

### Building and Testing

```powershell
# Install dependencies
npm install

# Compile TypeScript
npm run compile

# Watch mode (auto-recompile)
npm run watch

# Lint code
npm run lint

# Package for testing
npm run package
```

### Testing Locally

**Option 1: F5 Extension Development Host**
1. Open extension folder in VS Code
2. Press F5 (opens Extension Development Host)
3. Test in the new window

**Option 2: Install VSIX**
1. `npm run package` to create VSIX
2. `Ctrl+Shift+P` → "Install from VSIX"
3. Select the generated `.vsix` file

---

## Publishing Workflow

### Automated Publishing (Preferred)

1. **Add a changeset** for user-visible changes.
2. **Dispatch the unified release workflow**:
   ```powershell
   gh workflow run release.yml -f version_bump=patch
   ```
3. **Unified GitHub Actions workflow handles the rest**

### Manual Publishing (Emergency Only)

If automated workflow fails:

```powershell
# Login to marketplace
npx @vscode/vsce login <publisher-name>

# Publish
npx @vscode/vsce publish
```

---

## Common Mistakes to Avoid

### ❌ Don't Do This

1. **Don't manually edit package.json version**
   - Workflow updates it automatically from the release input
   
2. **Don't document a separate runtime install**
   - Extension uses its bundled self-contained MCP Server
   
3. **Don't edit CHANGELOG.md directly**
   - Add a changeset; the workflow generates the marketplace changelog
   
4. **Don't commit with outdated version references**
   - Check README.md, DEVELOPMENT.md for correct command syntax

### ✅ Do This

1. **Add changesets** for user-visible changes
2. **Keep bundled-server packaging references accurate**
3. **Let the workflow manage versions** from dispatch inputs
4. **Test locally** before dispatching a release
5. **Update README.md** when features change

---

## Key Principles

1. **Changesets drive CHANGELOG.md** - Never maintain release sections manually
2. **Workflow manages versions** - Don't manually edit package.json
3. **Bundled runtime** - The extension registers its packaged self-contained server
4. **Marketplace accuracy** - README.md and CHANGELOG.md must be current
5. **Test before release** - Use F5 or local VSIX install

---

## References

- **Main Extension Docs**: [vscode-extension/DEVELOPMENT.md](../../DEVELOPMENT.md)
- **Marketplace Publishing**: [vscode-extension/MARKETPLACE-PUBLISHING.md](../../MARKETPLACE-PUBLISHING.md)
- **Release Workflow**: [.github/workflows/release.yml](../../../.github/workflows/release.yml)
- **VS Code Extension API**: https://code.visualstudio.com/api
