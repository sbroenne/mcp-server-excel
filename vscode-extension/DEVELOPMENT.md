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
├── skills/                   # Agent skills (copied during build)
│   ├── excel-mcp/            # MCP server skill
│   │   └── SKILL.md
│   ├── excel-cli/            # CLI skill
│   │   └── SKILL.md
│   └── shared/               # Shared reference docs
│       └── *.md
└── excelmcp-1.0.0.vsix      # Packaged extension
```

## Key Implementation Details

### MCP Server Registration

The extension uses VS Code's `mcpServerDefinitionProvider` contribution point:

```typescript
vscode.lm.registerMcpServerDefinitionProvider('excelmcp', {
  provideMcpServerDefinitions: async () => {
    const serverPath = path.join(context.extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');
    return [
      new vscode.McpStdioServerDefinition(
        'Excel MCP Server',
        serverPath,
        [],
        {} // Optional environment variables
      )
    ];
  }
})
```

### Agent Skills Registration

The extension uses VS Code's `chatSkills` contribution point in `package.json` to declaratively register agent skills:

```json
"chatSkills": [
  { "name": "excel-mcp", "path": "./skills/excel-mcp/SKILL.md" },
  { "name": "excel-cli", "path": "./skills/excel-cli/SKILL.md" }
]
```

Skills are automatically available to GitHub Copilot when the extension is active — no file-copying needed.

### Activation

- **Activation Event**: `onStartupFinished` - Extension loads when VS Code starts
- **Welcome Message**: Shows once on first activation
- **State Management**: Uses `context.globalState` to track welcome message

### Dependencies

- **Runtime**: None - Extension bundles self-contained executables (MCP Server + CLI)
- **Dev Dependencies**:
  - `@types/vscode@^1.106.0` - VS Code API types
  - `@types/node@^22.0.0` - Node.js types
  - `typescript@^5.9.0` - TypeScript compiler
  - `@vscode/vsce@^3.0.0` - Extension packaging tool
  - `eslint` + `typescript-eslint` - Code quality

## Building

```powershell
npm install          # Install dependencies
npm run compile      # Compile TypeScript
npm run watch        # Watch mode for development
npm run lint         # Run ESLint
npm run package      # Create VSIX package
```

## Building Bundled Executables

The extension includes self-contained MCP server and CLI executables. To update them:

```powershell
# Build MCP server as self-contained single-file exe
cd d:\source\mcp-server-excel
dotnet publish src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false -p:PublishReadyToRun=false -p:NuGetAudit=false -o vscode-extension/bin

# Build CLI as self-contained single-file exe
dotnet publish src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false -p:PublishReadyToRun=false -p:NuGetAudit=false -o vscode-extension/bin

# Or use the npm script which builds both
npm run build:all

# Verify the executables work
vscode-extension/bin/Sbroenne.ExcelMcp.McpServer.exe --version
vscode-extension/bin/excelcli.exe --version
```

This creates self-contained executables with the .NET runtime and all dependencies included. No .NET SDK or runtime installation needed on end-user machines.

## Testing

### Prerequisites for Testing

The extension uses bundled self-contained executables. For development testing:

```powershell
# Build both executables (matches production)
npm run build:all

# Verify bundled executables work
vscode-extension/bin/Sbroenne.ExcelMcp.McpServer.exe --version
vscode-extension/bin/excelcli.exe --version
```

**Why this approach**: The extension bundles self-contained MCP server and CLI executables. No .NET runtime or SDK needed on the target machine.

### Manual Testing

1. **Build the extension**:
   ```powershell
   npm run compile
   ```

2. **Press F5 in VS Code** (opens Extension Development Host)

3. **Check the Debug Console** for activation logs:
   - ✅ `ExcelMcp extension is now active`
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
   ```powershell
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

The extension is published with every unified repository release:

```powershell
# Release all components with the same calculated version
gh workflow run release.yml -f version_bump=patch
```

The GitHub Actions workflow will automatically:
- ✅ **Calculate version** from the latest tag and dispatch input
- ✅ **Update package.json version** using `npm version` (no manual editing needed)
- ✅ **Compile changesets** into `CHANGELOG.md` and release notes
- ✅ **Build and package the extension**
- ✅ **Publish to VS Code Marketplace** (if `VSCE_TOKEN` secret is configured)
- ✅ **Build all other components** (MCP Server, CLI, MCPB)
- ✅ **Create unified GitHub release** with all artifacts

**Important**: The workflow manages version numbers. Do not manually update
`package.json` before release.

See [MARKETPLACE-PUBLISHING.md](MARKETPLACE-PUBLISHING.md) for setup instructions.

## Changelog Maintenance

Never edit the root `CHANGELOG.md` manually. Add a changeset with
`npx changeset` for every user-visible change. The unified release workflow
consumes pending fragments, generates the changelog entry, and copies the
generated changelog into the extension package.

### Manual Publishing

#### VS Code Marketplace

1. **Create publisher account**: https://marketplace.visualstudio.com/manage
2. **Generate PAT**: https://dev.azure.com (Marketplace Manage scope)
3. **Login**: `npx @vscode/vsce login <publisher>`
4. **Publish**: `npx @vscode/vsce publish`

#### GitHub Releases Only

To create a GitHub release without marketplace publishing:

```powershell
cd vscode-extension
npm run package
# Upload the .vsix file manually to GitHub releases
```

## Versioning

**Automatic Version Management** (Recommended):
The unified release workflow automatically calculates version numbers from the latest git tag:

1. Go to **Actions** → **Release All Components** → **Run workflow**
2. Select version bump type (patch/minor/major) or enter a custom version

The workflow will:
- Calculate the next version from the latest git tag
- Update `package.json` version for VS Code extension
- Update all component versions (MCP Server, CLI, MCPB manifest)
- Create git tag and unified GitHub release with all artifacts

**Local Version Testing**:
If packaging needs a temporary local version:

```powershell
npm version patch   # Bumps 1.0.0 → 1.0.1
npm version minor   # Bumps 1.0.0 → 1.1.0
npm version major   # Bumps 1.0.0 → 2.0.0
```

Follow Semantic Versioning (SemVer):
- **Major**: Breaking changes
- **Minor**: New features
- **Patch**: Bug fixes

**Important**: Don't commit local version changes. Releases use workflow inputs.

## Maintenance

### Updating Dependencies

```powershell
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
- Ensure bundled executable exists in `bin/Sbroenne.ExcelMcp.McpServer.exe`
- Run `npm run build:all` to build both MCP server and CLI executables
- Verify bundled executable runs: `bin/Sbroenne.ExcelMcp.McpServer.exe --version`

**CLI not found**
- Ensure `bin/excelcli.exe` exists
- Run `npm run build:all` to build both executables

## Extension Size 

Current size: **~68-70 MB** (includes bundled self-contained MCP server and CLI executables)

The extension includes:
- Main extension code (~10 KB)
- Bundled self-contained MCP server (~118 MB uncompressed, ~34 MB compressed)
- Bundled self-contained CLI (~115 MB uncompressed, ~34 MB compressed)
- Agent Skills (~130 KB for both excel-mcp and excel-cli)

Benefits of self-contained bundled approach:
- ✅ Zero-setup installation (no .NET runtime or SDK required)
- ✅ Version compatibility guaranteed (extension includes matching MCP server + CLI)
- ✅ Works offline after installation
- ✅ No dependency on dotnet tool installations
- ✅ CLI available directly for terminal-based automation

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
