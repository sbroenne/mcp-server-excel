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

## CHANGELOG.md Maintenance (CRITICAL)

### How to Maintain CHANGELOG.md

**Rule:** The root CHANGELOG.md should always have a **top entry ready for the next release**. The release workflow will automatically update the version number and date.

### Workflow Process

1. **You maintain**: Keep root CHANGELOG.md updated with changes as you make them
2. **Version number must match tag**: Use the version you'll tag (e.g., `1.5.7`)
3. **Workflow extracts automatically**: When you push tag `v1.5.7`, the unified workflow:
 - Extracts the section for that version from CHANGELOG.md for release notes
 - Updates package.json version to `1.5.7`
 - Builds and releases ALL components (MCP Server, CLI, VS Code Extension, MCPB)

### Example Workflow

**During Development** (root CHANGELOG.md):
```markdown
# Changelog

## [Unreleased]

## [1.5.7] - 2025-01-21

### Added
- New feature A
- New feature B

### Fixed
- Bug fix C

## [1.5.6] - 2025-01-20

### Added
- Initial slicer support
```

**After Pushing Tag** `v1.5.7`:
- Workflow extracts the `[1.5.7]` section for GitHub Release notes
- All 4 components are built and released with version 1.5.7
```

### Best Practice

**After each release, add a new top section**:
```markdown
# Change Log

## [1.0.0] - 2025-10-29

### Added
- Preparing for next release
- Add changes here as you develop

## [1.1.0] - 2025-10-30

### Added
- Previous release
```

This ensures CHANGELOG is always ready for the next release.

### Format Guidelines

Follow [Keep a Changelog](https://keepachangelog.com/) format:

- **Added**: New features
- **Changed**: Changes in existing functionality
- **Deprecated**: Soon-to-be removed features
- **Removed**: Removed features
- **Fixed**: Bug fixes
- **Security**: Security fixes

---

## Version Management

### Automatic Version Management (Unified Release Workflow)

**DO NOT manually edit package.json version** - The unified release workflow handles this:

```powershell
# Create and push tag - workflow releases ALL components with same version
git tag v1.5.7
git push origin v1.5.7
```

Unified workflow automatically:
1. Extracts version from tag (`v1.5.7` → `1.5.7`)
2. Updates `package.json` version using `npm version`
3. Extracts changelog section for release notes
4. Builds and packages VS Code extension
5. Builds all other components (MCP Server, CLI, MCPB)
6. Publishes to VS Code Marketplace and NuGet
7. Creates unified GitHub release with all artifacts

### Local Testing (Manual Version Bump)

For local testing only, use npm version commands:

```powershell
npm version patch # 1.0.0 → 1.0.1
npm version minor # 1.0.0 → 1.1.0
npm version major # 1.0.0 → 2.0.0
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

- **README.md** - Keep up-to-date with accurate commands and features
- **CHANGELOG.md** - Maintain version history
- **package.json** - Ensure metadata is accurate
- **icon.png** - High-quality 128x128px PNG

---

## Extension Commands

### Correct Command Syntax

**The extension uses**: `dotnet tool run mcp-excel`

**NOT**: `dnx Sbroenne.ExcelMcp.McpServer --yes` (this is incorrect)

### Where Commands Are Referenced

Check these files when updating command syntax:

1. **src/extension.ts** - Actual command executed
2. **README.md** - Documentation shown in marketplace
3. **DEVELOPMENT.md** - Developer notes
4. **INSTALL.md** - Installation guide (if applicable)

### Verification

Before committing, search for outdated command references:

```powershell
# Search for incorrect dnx references
grep -r "dnx" vscode-extension/

# Should only find references in documentation explaining the NuGet approach
# Actual command should be: dotnet tool run mcp-excel
```

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

1. **Update root CHANGELOG.md** with new features/fixes for all components
2. **Create and push version tag** (releases ALL components):
 ```powershell
 git tag v1.5.7
 git push origin v1.5.7
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

### Don't Do This

1. **Don't manually edit package.json version** before tagging
 - Workflow updates it automatically from tag
 
2. **Don't use dnx commands in documentation**
 - Extension uses `dotnet tool run mcp-excel`
 
3. **Don't forget to update CHANGELOG.md**
 - Marketplace shows changelog - keep it current
 
4. **Don't commit with outdated version references**
 - Check README.md, DEVELOPMENT.md for correct command syntax

### Do This

1. **Keep CHANGELOG.md updated** as you develop
2. **Use correct command syntax** (`dotnet tool run mcp-excel`)
3. **Let workflow manage versions** via git tags
4. **Test locally** before pushing tags
5. **Update README.md** when features change

---

## Key Principles

1. **CHANGELOG.md is always ready** - Top entry is for next release
2. **Workflow manages versions** - Don't manually edit package.json
3. **Correct command syntax** - `dotnet tool run mcp-excel` (not dnx)
4. **Marketplace accuracy** - README.md and CHANGELOG.md must be current
5. **Test before release** - Use F5 or local VSIX install

---

## References

- **Main Extension Docs**: [vscode-extension/DEVELOPMENT.md](../../DEVELOPMENT.md)
- **Marketplace Publishing**: [vscode-extension/MARKETPLACE-PUBLISHING.md](../../MARKETPLACE-PUBLISHING.md)
- **Release Workflow**: [.github/workflows/release-vscode-extension.yml](../../../.github/workflows/release-vscode-extension.yml)
- **VS Code Extension API**: https://code.visualstudio.com/api
