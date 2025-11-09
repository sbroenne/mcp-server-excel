# ExcelMcp Release Strategy

This document outlines the build and release processes for ExcelMcp components: MCP Server & CLI (unified), and VS Code Extension.

## Release Workflows

### 1. MCP Server & CLI Unified Releases (`v*` tags)

**Workflow**: `.github/workflows/release-mcp-server.yml`
**Trigger**: Tags starting with `v` (e.g., `v1.0.0`)

**Features**:

- Builds and packages **both MCP Server and CLI together**
- Publishes both to NuGet as .NET tools (using OIDC trusted publishing)
- Creates unified GitHub release with both packages
- Synchronized versioning ensures compatibility
- Single workflow handles both NuGet packages and GitHub release

**Release Artifacts**:

- `ExcelMcp-MCP-Server-{version}-windows.zip` - MCP Server binary package
- `ExcelMcp-CLI-{version}-windows.zip` - CLI binary package
- NuGet package: `Sbroenne.ExcelMcp.McpServer` on NuGet.org
- NuGet package: `Sbroenne.ExcelMcp.CLI` on NuGet.org (as .NET global tool)
- Unified release notes covering both packages

**Publishing Method**:
- Uses OIDC (OpenID Connect) trusted publishing for secure NuGet authentication
- No API keys stored in secrets - authentication via GitHub identity
- Publishes both packages to NuGet.org within same workflow
- CLI configured as .NET global tool (PackAsTool=true)

**Use Cases**:

- **MCP Server**: AI assistant integration (GitHub Copilot, Claude, ChatGPT), conversational Excel workflows
- **CLI**: Direct Excel automation scripts, CI/CD pipeline integration, command-line operations
- **Both**: Unified version ensures compatibility, simplified dependency management

**Why Unified?**
- MCP Server and CLI share Core/ComInterop libraries as internal dependencies
- Synchronized versioning prevents compatibility issues
- Simplified release process (one tag, one workflow)
- Users can install either or both with matching versions

### 2. VS Code Extension Releases (`vscode-v*` tags)

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

### Versioning Strategy

**Unified Packages:**
- **MCP Server & CLI**: Always released together with same version number (e.g., v1.2.0)
- Single tag triggers both NuGet packages
- Ensures compatibility between MCP Server and CLI
- Core and ComInterop are internal dependencies (not separately released to NuGet)

**Independent Package:**
- **VS Code Extension**: Independent version numbers (e.g., vscode-v1.0.5)

### Development Strategy

- **Core & ComInterop**: Internal libraries, not separately published to NuGet
- **MCP Server & CLI**: Consumer-facing packages, always co-released
  - **MCP Server**: AI integration features, conversational interfaces
  - **CLI**: Direct automation, command completeness, CI/CD integration
- **VS Code Extension**: Focus on developer experience, VS Code integration, MCP management

## Release Process Examples

### Releasing MCP Server & CLI Together (Standard)

```bash
# MCP Server and CLI are always released together

# Create and push unified release tag
git tag v1.2.0
git push origin v1.2.0

# This triggers release-mcp-server.yml which:
# - Builds both MCP Server and CLI
# - Publishes both to NuGet using OIDC trusted publishing
# - Creates GitHub release with both ZIP packages
# - Unified release notes covering both packages
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
- **src/ExcelMcp.CLI/README.md**: CLI focused (direct automation)
- **vscode-extension/README.md**: VS Code Extension focused (developer experience)
- **Release Notes**: Unified for MCP Server + CLI, separate for VS Code Extension

### Cross-References

- Each tool's documentation references the others
- Clear navigation between MCP, CLI, and VS Code Extension docs
- Unified project branding while maintaining component clarity

## Benefits of This Approach

1. **Unified Releases**: MCP Server and CLI always compatible (same version)
2. **Simplified Process**: One tag triggers both packages
3. **Reduced Complexity**: No need to coordinate separate Core/ComInterop releases
4. **Focused Documentation**: Release notes match user intent
5. **Clear Separation**: MCP for AI workflows, CLI for automation, VS Code for IDE
6. **Flexibility**: VS Code Extension evolves independently
7. **NuGet Ecosystem**: Both tools available via NuGet package manager

## Tag Patterns

- `v*`: MCP Server and CLI together (unified release)
- `vscode-v*`: VS Code Extension only (VSIX)

This approach provides simplified release management while maintaining the integrated ExcelMcp ecosystem.
