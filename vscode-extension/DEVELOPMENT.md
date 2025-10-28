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
        'dnx',
        ['Sbroenne.ExcelMcp.McpServer', '--yes'],
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

- **Runtime**: None - Uses `dnx` command from .NET SDK
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

### Marketplace Publishing

1. **Create publisher account**: https://marketplace.visualstudio.com/
2. **Generate PAT**: https://dev.azure.com
3. **Login**: `npx @vscode/vsce login <publisher>`
4. **Publish**: `npx @vscode/vsce publish`

### GitHub Releases

1. **Tag version**: `git tag v1.0.0`
2. **Push tag**: `git push --tags`
3. **Create release** on GitHub
4. **Upload VSIX** as release asset

## Publishing

### Automated Publishing (Recommended)

The extension is automatically published to both marketplaces when a version tag is pushed:

```bash
# 1. Update version in package.json and CHANGELOG.md
npm version patch  # or minor, or major

# 2. Commit changes
git add .
git commit -m "Bump version to X.Y.Z"

# 3. Create and push tag
git tag vscode-vX.Y.Z
git push && git push --tags
```

The GitHub Actions workflow will:
- Build and package the extension
- Publish to VS Code Marketplace (if `VSCE_TOKEN` secret is configured)
- Publish to Open VSX Registry (if `OPEN_VSX_TOKEN` secret is configured)
- Create GitHub release with VSIX file

See [MARKETPLACE-PUBLISHING.md](MARKETPLACE-PUBLISHING.md) for setup instructions.

### Manual Publishing

#### VS Code Marketplace

1. **Create publisher account**: https://marketplace.visualstudio.com/manage
2. **Generate PAT**: https://dev.azure.com (Marketplace Manage scope)
3. **Login**: `npx @vscode/vsce login <publisher>`
4. **Publish**: `npx @vscode/vsce publish`

#### Open VSX Registry

1. **Create account**: https://open-vsx.org
2. **Generate token**: https://open-vsx.org/user-settings/tokens
3. **Publish**: `npx ovsx publish -p <token>`

#### GitHub Releases Only

To create a GitHub release without marketplace publishing:

```bash
cd vscode-extension
npm run package
# Upload the .vsix file manually to GitHub releases
```

## Versioning

Follow Semantic Versioning (SemVer):
- **Major**: Breaking changes
- **Minor**: New features
- **Patch**: Bug fixes

Update version in:
- `package.json`
- `CHANGELOG.md`

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
- Ensure `dnx` command is available
- Check .NET 10 SDK is installed
- Verify NuGet package name is correct

## Extension Size Optimization

Current size: **9 KB** (very small!)

Ways to keep it small:
- ✅ Use `--no-dependencies` when packaging (only include compiled code)
- ✅ Use `.vscodeignore` to exclude source files
- ✅ No runtime dependencies (uses dnx)
- ✅ Minimal icon size (1 KB)

## Future Enhancements

Potential improvements:
- [ ] Add configuration options for MCP server
- [ ] Status bar item showing server status
- [ ] Commands to restart/reload MCP server
- [ ] Settings for custom dnx arguments
- [ ] Telemetry for usage insights
- [ ] Automatic update notifications

## References

- [VS Code Extension API](https://code.visualstudio.com/api)
- [MCP Documentation](https://modelcontextprotocol.io/)
- [VS Code Extension Samples](https://github.com/microsoft/vscode-extension-samples)
- [Publishing Extensions](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
