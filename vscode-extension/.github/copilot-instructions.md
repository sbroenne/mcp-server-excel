# GitHub Copilot Instructions - ExcelMcp VS Code Extension

> **TypeScript VS Code Extension for Excel MCP Server Integration**

## Overview

This VS Code extension provides one-click installation of the ExcelMcp MCP server, enabling AI assistants like GitHub Copilot to automate Microsoft Excel through natural language.

**Key Technologies:**
- TypeScript with ES2022 target
- VS Code Extension API (^1.106.0)
- MCP Server Definition Provider API
- Bundled .NET 8 self-contained executable

---

## Project Structure

```
vscode-extension/
├── src/extension.ts       # Extension entry point (activation, MCP registration)
├── package.json           # Extension manifest (metadata, version, dependencies)
├── tsconfig.json          # TypeScript configuration
├── eslint.config.mjs      # ESLint rules (flat config)
├── README.md              # Marketplace description
├── CHANGELOG.md           # Version history
├── bin/                   # Bundled MCP server executable
└── out/                   # Compiled JavaScript output
```

---

## Coding Standards

### TypeScript
- Use strict mode (`"strict": true` in tsconfig.json)
- Use ES2022 features and Node16 module resolution
- Prefer `const` over `let`; avoid `var`
- Use explicit type annotations for function parameters and return types

### Naming Conventions
- **Functions**: camelCase (`ensureDotNetRuntime`, `showWelcomeMessage`)
- **Constants**: UPPER_SNAKE_CASE for true constants, camelCase for const variables
- **Imports**: camelCase or PascalCase

### VS Code API Patterns
```typescript
// Extension activation
export async function activate(context: vscode.ExtensionContext) {
    // Register disposables with context.subscriptions
    context.subscriptions.push(
        vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', provider)
    );
}

// Error handling
try {
    await asyncOperation();
} catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    vscode.window.showErrorMessage(`ExcelMcp: ${errorMessage}`);
}
```

---

## Development Commands

```bash
# Install dependencies
npm install

# Compile TypeScript
npm run compile

# Watch mode (auto-recompile on changes)
npm run watch

# Lint code
npm run lint

# Package extension (includes bundled MCP server)
npm run package

# Build bundled MCP server executable
npm run build:mcp-server
```

---

## Testing

### Local Testing (F5 Method)
1. Open extension folder in VS Code
2. Press F5 (opens Extension Development Host)
3. Check Debug Console for activation logs:
   - ✅ `ExcelMcp extension is now active`
   - ✅ `ExcelMcp: .NET runtime setup completed`

### VSIX Testing
1. Run `npm run package` to create VSIX
2. `Ctrl+Shift+P` → "Install from VSIX"
3. Verify extension loads and MCP server is registered

---

## Version Management

**DO NOT manually edit `package.json` version** - The release workflow handles this:

```bash
# Create and push tag - workflow does everything
git tag vscode-v1.2.3
git push origin vscode-v1.2.3
```

The workflow automatically:
1. Extracts version from tag
2. Updates package.json
3. Updates CHANGELOG.md with release date
4. Publishes to VS Code Marketplace

---

## CHANGELOG.md Maintenance

Keep a **top entry ready for next release** with placeholder version:

```markdown
## [X.Y.Z] - YYYY-MM-DD

### Added
- New feature description

### Fixed
- Bug fix description
```

The release workflow replaces the version and date automatically.

---

## Common Mistakes to Avoid

| ❌ Don't | ✅ Do |
|----------|-------|
| Manually edit package.json version | Use git tags for releases |
| Forget to add disposables to subscriptions | `context.subscriptions.push(...)` |
| Use synchronous file operations | Use `vscode.workspace.fs` async API |
| Ignore error types | Check `error instanceof Error` |
| Skip lint before commit | Run `npm run lint` |

---

## Extension Manifest (package.json)

**Critical fields:**
- `displayName`: Title shown in marketplace
- `description`: Brief subtitle
- `engines.vscode`: Minimum VS Code version
- `activationEvents`: Use `onStartupFinished` for background activation
- `extensionDependencies`: Include required extensions
- `contributes.mcpServerDefinitionProviders`: MCP server registration

---

## MCP Server Integration

The extension registers an MCP server using the bundled executable:

```typescript
const extensionPath = context.extensionPath;
const mcpServerPath = path.join(extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe');

new vscode.McpStdioServerDefinition(
    'ExcelMcp - Excel Automation',
    mcpServerPath,
    [],
    {} // Environment variables
)
```

**Note:** The MCP server is bundled as a self-contained .NET 8 executable (~41 MB).

---

## References

- [VS Code Extension API](https://code.visualstudio.com/api)
- [MCP Server Definition Provider](https://code.visualstudio.com/api/references/vscode-api#McpServerDefinitionProvider)
- [Publishing Extensions](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
- [Main Project Docs](https://github.com/sbroenne/mcp-server-excel)
