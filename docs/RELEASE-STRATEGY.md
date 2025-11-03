# ExcelMcp Release Strategy

This document outlines the separate build and release processes for ExcelMcp components: MCP Server, CLI, and VS Code Extension.

## Release Workflows

### 1. MCP Server Releases (`mcp-v*` tags)

**Workflow**: `.github/workflows/release-mcp-server.yml`
**Trigger**: Tags starting with `mcp-v` (e.g., `mcp-v1.0.0`)

**Features**:

- Builds and packages only the MCP Server
- Publishes to NuGet as a .NET tool (using OIDC trusted publishing)
- Creates GitHub release with MCP-focused documentation
- Optimized for AI assistant integration
- Single workflow handles both NuGet publishing and GitHub release

**Release Artifacts**:

- `ExcelMcp-MCP-Server-{version}-windows.zip` - Binary package
- NuGet package: `Sbroenne.ExcelMcp.McpServer` on NuGet.org
- Installation guide focused on MCP usage

**Publishing Method**:
- Uses OIDC (OpenID Connect) trusted publishing for secure NuGet authentication
- No API keys stored in secrets - authentication via GitHub identity
- Publishes to NuGet.org within the same workflow that creates the GitHub release

**Use Cases**:

- AI assistant integration (GitHub Copilot, Claude, ChatGPT)
- Conversational Excel development workflows
- Model Context Protocol implementations

### 2. CLI Releases (`cli-v*` tags)

**Workflow**: `.github/workflows/release-cli.yml`
**Trigger**: Tags starting with `cli-v` (e.g., `cli-v1.0.0`)

**Features**:

- Builds and packages only the CLI tool
- Creates standalone CLI distribution
- Focused on direct automation workflows
- No NuGet publishing (binary-only distribution)

**Release Artifacts**:

- `ExcelMcp-CLI-{version}-windows.zip` - Complete CLI package
- Includes all 40+ commands documentation
- Quick start guide for CLI usage

**Use Cases**:

- Direct Excel automation scripts
- CI/CD pipeline integration
- Development workflows and testing
- Command-line Excel operations

### 3. VS Code Extension Releases (`vscode-v*` tags)

**Workflow**: `.github/workflows/release-vscode-extension.yml`
**Trigger**: Tags starting with `vscode-v` (e.g., `vscode-v1.0.0`)

**Features**:

- Builds and packages VS Code extension
- Creates VSIX package for VS Code Marketplace
- Focused on VS Code integration
- No NuGet publishing (VSIX distribution only)

**Release Artifacts**:

- `excel-mcp-{version}.vsix` - VS Code Extension package
- Installation guide for VS Code extension
- Extension marketplace listing updates

**Use Cases**:

- VS Code integration for Excel development
- MCP server management within VS Code
- Developer experience improvements
- One-click MCP server configuration

## Version Management

### Independent Versioning

- **MCP Server**: Independent version numbers (e.g., mcp-v1.2.0)
- **CLI**: Independent version numbers (e.g., cli-v2.1.0)
- **VS Code Extension**: Independent version numbers (e.g., vscode-v1.0.5)

### Development Strategy

- **MCP Server**: Focus on AI integration features, conversational interfaces
- **CLI**: Focus on automation efficiency, command completeness, CI/CD integration
- **VS Code Extension**: Focus on developer experience, VS Code integration, MCP management

## Release Process Examples

### Releasing MCP Server Only

```bash
# Create and push MCP server release tag
git tag mcp-v1.3.0
git push origin mcp-v1.3.0

# This triggers release-mcp-server.yml which:
# - Builds MCP server
# - Publishes to NuGet using OIDC trusted publishing
# - Creates GitHub release with MCP-focused docs and binary ZIP
```

### Releasing CLI Only

```bash
# Create and push CLI release tag
git tag cli-v2.2.0
git push origin cli-v2.2.0

# This triggers release-cli.yml which:
# - Builds CLI only
# - Creates binary distribution (no NuGet publishing)
# - Creates GitHub release with CLI-focused docs
```

### Releasing VS Code Extension Only

```bash
# Create and push VS Code extension release tag
git tag vscode-v1.0.5
git push origin vscode-v1.0.5

# This triggers release-vscode-extension.yml which:
# - Builds VS Code extension
# - Creates VSIX package
# - Creates GitHub release with extension installation docs
```

## Documentation Strategy

### Separate Focus Areas

- **Main README.md**: MCP Server focused (AI assistant integration)
- **docs/CLI.md**: CLI focused (direct automation)
- **vscode-extension/README.md**: VS Code Extension focused (developer experience)
- **Release Notes**: Tailored to the specific component being released

### Cross-References

- Each tool's documentation references the others
- Clear navigation between MCP, CLI, and VS Code Extension docs
- Unified project branding while maintaining component clarity

## Benefits of This Approach

1. **Targeted Releases**: Users can get updates for just the tool they use
2. **Independent Development**: MCP, CLI, and VS Code Extension can evolve at different paces
3. **Focused Documentation**: Release notes and docs match user intent
4. **Reduced Package Size**: Users download only what they need
5. **Clear Separation**: MCP for AI workflows, CLI for automation, VS Code Extension for IDE integration
6. **Flexibility**: Each component released independently as needed

## Tag Patterns

- `mcp-v*`: MCP Server only
- `cli-v*`: CLI only  
- `vscode-v*`: VS Code Extension only

This approach provides maximum flexibility while maintaining the integrated ExcelMcp ecosystem.
