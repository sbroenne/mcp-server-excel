# VS Code Extension Development Notes

## Project Structure

```
vscode-extension/
├── src/
│   └── extension.ts          # Extension entry point
├── out/                       # Compiled JavaScript
│   ├── extension.js
│   └── extension.js.map
├── package.json               # Extension manifest
├── tsconfig.json             # TypeScript config
├── eslint.config.mjs         # Linting rules
├── README.md                 # Extension documentation
├── CHANGELOG.md              # Version history
├── INSTALL.md                # Installation guide
├── LICENSE                   # MIT License
├── icon.png                  # 128x128 extension icon
├── icon.svg                  # SVG source
└── excelmcp-1.0.0.vsix      # Packaged extension
```

## Key Implementation Details

### MCP Server Registration

The extension uses VS Code's `mcpServerDefinitionProvider` contribution point:

```typescript
vscode.lm.registerMcpServerDefinitionProvider('excelmcp', {
  provideMcpServerDefinitions: async () => {
    return [
      new vscode.McpStdioServerDefinition(
        'ExcelMcp - Excel Automation',
        'dotnet',
        ['tool', 'run', 'mcp-excel'],
        {} // Optional environment variables
      )
    ];
  }
})
```

### Activation

- **Activation Event**: `onStartupFinished` - Extension loads when VS Code starts
- **Welcome Message**: Shows once on first activation
- **State Management**: Uses `context.globalState` to track welcome message

### Dependencies

- **Runtime**: None - Uses `dotnet tool run` command from .NET SDK
- **Dev Dependencies**:
  - `@types/vscode@^1.105.0` - VS Code API types
  - `@types/node@^22.0.0` - Node.js types
  - `typescript@^5.9.0` - TypeScript compiler
  - `@vscode/vsce@^3.0.0` - Extension packaging tool
  - `eslint` + `typescript-eslint` - Code quality

## Building

```bash
npm install          # Install dependencies
npm run compile      # Compile TypeScript
npm run watch        # Watch mode for development
npm run lint         # Run ESLint
npm run package      # Create VSIX package
```

## Testing

### Manual Testing

1. **Build the extension**:
   ```bash
   npm run compile
   ```

2. **Press F5 in VS Code** (opens Extension Development Host)

3. **In the Extension Development Host**:
   - Check if extension is loaded: Extensions panel
   - Check if MCP server is registered: Settings → MCP
   - Ask GitHub Copilot to list Excel tools

### Package Testing

1. **Package the extension**:
   ```bash
   npm run package
   ```

2. **Install from VSIX**:
   - `Ctrl+Shift+P` → "Install from VSIX"
   - Select `excelmcp-1.0.0.vsix`

3. **Verify**:
   - Extension appears in Extensions panel
   - Welcome message shows on first activation
   - GitHub Copilot can access Excel tools

## Publishing

### Automated Publishing (Recommended)

The extension is automatically published to the VS Code Marketplace when a version tag is pushed:

```bash
# 1. Create and push tag (workflow updates version automatically)
git tag vscode-vX.Y.Z
git push --tags
```

The GitHub Actions workflow will automatically:
- ✅ **Extract version from tag** (e.g., `vscode-v1.0.0` → `1.0.0`)
- ✅ **Update package.json version** using `npm version` (no manual editing needed)
- ✅ **Update CHANGELOG.md** with release date
- ✅ **Build and package the extension**
- ✅ **Publish to VS Code Marketplace** (if `VSCE_TOKEN` secret is configured)
- ✅ **Create GitHub release** with VSIX file

**Important**: The workflow manages version numbers - you don't need to manually update `package.json` before tagging.

See [MARKETPLACE-PUBLISHING.md](MARKETPLACE-PUBLISHING.md) for setup instructions.

## CHANGELOG Maintenance

### How to Maintain CHANGELOG.md

The CHANGELOG.md file should always have a **top entry ready for the next release**. The release workflow will automatically update the version number and date.

**Before Release**:
```markdown
## [1.0.0] - 2025-10-29

### Added
- New feature A
- New feature B

### Fixed
- Bug fix C
```

**After Release** (workflow automatically updates):
```markdown
## [1.1.0] - 2025-10-30

### Added
- New feature A
- New feature B

### Fixed
- Bug fix C
```

### Workflow Process

1. **You maintain**: Keep CHANGELOG.md updated with changes, but version number can be any placeholder
2. **Workflow updates**: When you push tag `vscode-v1.1.0`, the workflow replaces the first version number with `1.1.0` and updates the date

### Best Practice

**After each release, add a new top section for the next version**:

```markdown
# Change Log

## [1.0.0] - 2025-10-29

### Added
- Prepare for next release
- Add changes here as you make them

## [1.0.0] - 2025-10-29

### Added
- Initial release
...
```

This way, the CHANGELOG is always ready, and the workflow just updates the version/date.

### Format

Follow [Keep a Changelog](https://keepachangelog.com/) format:
- **Added**: New features
- **Changed**: Changes in existing functionality
- **Deprecated**: Soon-to-be removed features
- **Removed**: Removed features
- **Fixed**: Bug fixes
- **Security**: Security fixes

### Manual Publishing

#### VS Code Marketplace

1. **Create publisher account**: https://marketplace.visualstudio.com/manage
2. **Generate PAT**: https://dev.azure.com (Marketplace Manage scope)
3. **Login**: `npx @vscode/vsce login <publisher>`
4. **Publish**: `npx @vscode/vsce publish`

#### GitHub Releases Only

To create a GitHub release without marketplace publishing:

```bash
cd vscode-extension
npm run package
# Upload the .vsix file manually to GitHub releases
```

## Versioning

**Automatic Version Management** (Recommended):
The release workflow automatically updates version numbers from git tags:

```bash
# Just create and push the tag - workflow does the rest
git tag vscode-v1.2.3
git push --tags
```

The workflow will:
- Extract version from tag (`vscode-v1.2.3` → `1.2.3`)
- Update `package.json` version
- Update `CHANGELOG.md` with release date

**Manual Version Updates** (if needed):
If you need to update the version locally before tagging:

```bash
npm version patch   # Bumps 1.0.0 → 1.0.1
npm version minor   # Bumps 1.0.0 → 1.1.0
npm version major   # Bumps 1.0.0 → 2.0.0
```

Follow Semantic Versioning (SemVer):
- **Major**: Breaking changes
- **Minor**: New features
- **Patch**: Bug fixes

**Important**: Don't manually edit version numbers in `package.json` - use either git tags (for releases) or `npm version` commands (for local testing).

## Maintenance

### Updating Dependencies

```bash
npm outdated                    # Check for updates
npm update                      # Update minor/patch
npm install @types/vscode@latest --save-dev  # Update major
```

### VS Code API Updates

When VS Code releases new API features:
1. Update `engines.vscode` in package.json
2. Update `@types/vscode` to matching version
3. Test extension compatibility
4. Update CHANGELOG

## Troubleshooting

### Build Issues

**Error: "Cannot find module 'vscode'"**
- Run `npm install`

**Error: "TypeScript compile errors"**
- Check `tsconfig.json` settings
- Verify VS Code types version matches engines.vscode

### Packaging Issues

**Error: "LICENSE not found"**
- Ensure LICENSE file exists in extension root

**Error: "engines.vscode mismatch"**
- Update package.json `engines.vscode` to match `@types/vscode` version

### Runtime Issues

**Extension not activating**
- Check `activationEvents` in package.json
- Verify extension ID matches registration

**MCP server not found**
- Ensure `dotnet tool run mcp-excel` command works
- Check .NET 8 Runtime is installed
- Verify NuGet package is available

## Extension Size Optimization

Current size: **9 KB** (very small!)

Ways to keep it small:
- ✅ Use `--no-dependencies` when packaging (only include compiled code)
- ✅ Use `.vscodeignore` to exclude source files
- ✅ No runtime dependencies (uses dotnet tool)
- ✅ Minimal icon size (1 KB)

## Future Enhancements

Potential improvements:
- [ ] Add configuration options for MCP server
- [ ] Status bar item showing server status
- [ ] Commands to restart/reload MCP server
- [ ] Settings for custom tool arguments
- [ ] Telemetry for usage insights
- [ ] Automatic update notifications

## References

- [VS Code Extension API](https://code.visualstudio.com/api)
- [MCP Documentation](https://modelcontextprotocol.io/)
- [VS Code Extension Samples](https://github.com/microsoft/vscode-extension-samples)
- [Publishing Extensions](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
