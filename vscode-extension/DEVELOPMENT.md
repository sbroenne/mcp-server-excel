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
        'Excel MCP Server',
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

## Building Bundled Executable

The extension includes a self-contained MCP server executable. To update it:

```bash
# 1. Navigate to MCP server project
cd d:\source\mcp-server-excel\src\ExcelMcp.McpServer

# 2. Publish self-contained executable for Windows x64
dotnet publish -c Release -r win-x64 --self-contained -o ../../vscode-extension/bin

# 3. Verify the executable works
../../vscode-extension/bin/Sbroenne.ExcelMcp.McpServer.exe --help
```

This creates a self-contained executable with all dependencies included.

## Testing

### Prerequisites for Testing

The extension uses a bundled MCP server executable. For development testing:

The extension uses a bundled MCP server executable. For development testing:

```bash
# Option 1: Use bundled executable (matches production)
# - Extension will use: extension-path/bin/Sbroenne.ExcelMcp.McpServer.exe
# - No additional setup needed

# Option 2: Test with local development version
# - Build and publish the MCP server as shown above
# - Extension automatically uses the bundled version

# Verify bundled executable works
cd vscode-extension
bin/Sbroenne.ExcelMcp.McpServer.exe --help
```

**Why this approach**: The extension uses a bundled MCP server executable. During development, you can use the local version or test with the bundled executable.

### Manual Testing

1. **Build the extension**:
   ```bash
   npm run compile
   ```

2. **Press F5 in VS Code** (opens Extension Development Host)

3. **Check the Debug Console** for activation logs:
   - ✅ `ExcelMcp extension is now active`
   - ✅ `ExcelMcp: .NET runtime available at ...`
   - ✅ `ExcelMcp: MCP server tool installation/update initiated`
   - ❌ NO errors about "Cannot read properties of undefined"

4. **In the Extension Development Host**:
   - Check if extension is loaded: Extensions panel
   - Check if MCP server is registered: Settings → MCP
   - Ask GitHub Copilot to list Excel tools

5. **Check Developer Tools Console** (Ctrl+Shift+I):
   - Go to Console tab
   - Look for "ExcelMcp:" messages
   - Verify no errors

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
# 1. Create and push tag (releases ALL components with same version)
git tag vX.Y.Z
git push --tags
```

The GitHub Actions workflow will automatically:
- ✅ **Extract version from tag** (e.g., `v1.5.7` → `1.5.7`)
- ✅ **Update package.json version** using `npm version` (no manual editing needed)
- ✅ **Update CHANGELOG.md** with release date
- ✅ **Build and package the extension**
- ✅ **Publish to VS Code Marketplace** (if `VSCE_TOKEN` secret is configured)
- ✅ **Build all other components** (MCP Server, CLI, MCPB)
- ✅ **Create unified GitHub release** with all artifacts

**Important**: The workflow manages version numbers - you don't need to manually update `package.json` before tagging. The unified release workflow (`.github/workflows/release.yml`) releases all components together.

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

1. **You maintain**: Keep root CHANGELOG.md updated with changes, but version number can be any placeholder
2. **Workflow updates**: When you push tag `v1.1.0`, the workflow extracts that version's section for release notes

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
The unified release workflow automatically updates version numbers from git tags:

```bash
# Just create and push the tag - workflow does the rest for ALL components
git tag v1.2.3
git push --tags
```

The workflow will:
- Extract version from tag (`v1.2.3` → `1.2.3`)
- Update `package.json` version for VS Code extension
- Update all component versions (MCP Server, CLI, MCPB manifest)
- Create unified GitHub release with all artifacts

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
- Ensure bundled executable exists in `bin/` directory
- Check .NET 10 Runtime is installed
- Verify bundled executable has all required dependencies

## Extension Size 

Current size: **~41 MB** (includes bundled MCP server executable)

The extension includes:
- Main extension code (~10 KB)
- Bundled .NET 10 self-contained MCP server (~41 MB)

Benefits of bundled approach:
- ✅ Zero-setup installation (no separate tool download required)
- ✅ Version compatibility guaranteed (extension includes matching MCP server)
- ✅ Works offline after installation
- ✅ No dependency on dotnet tool installations

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
