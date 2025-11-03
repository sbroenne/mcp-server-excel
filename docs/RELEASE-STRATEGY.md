# ExcelMcp Release Strategy

This document outlines the separate build and release processes for ExcelMcp components: MCP Server, CLI, and VS Code Extension.

## Release Workflows

### 1. ComInterop Library Releases (`cominterop-v*` tags)

**Workflow**: `.github/workflows/release-cominterop.yml`
**Trigger**: Tags starting with `cominterop-v` (e.g., `cominterop-v1.0.0`)

**Features**:

- Builds and packages only the ComInterop library
- Publishes to NuGet as a reusable library (using OIDC trusted publishing)
- Creates GitHub release with library-focused documentation
- Foundation layer for Excel COM automation
- Can be used independently in any .NET Excel automation project

**Release Artifacts**:

- NuGet package: `Sbroenne.ExcelMcp.ComInterop` on NuGet.org
- No binary ZIP (library only, consumed via NuGet)

**Publishing Method**:
- Uses OIDC (OpenID Connect) trusted publishing for secure NuGet authentication
- No API keys stored in secrets - authentication via GitHub identity
- Publishes to NuGet.org within the same workflow that creates the GitHub release

**Use Cases**:

- Low-level COM interop for Excel automation projects
- STA threading and OLE message filtering
- Excel session and batch operation management
- Reusable foundation for custom Excel automation tools

### 2. Core Library Releases (`core-v*` tags)

**Workflow**: `.github/workflows/release-core.yml`
**Trigger**: Tags starting with `core-v` (e.g., `core-v1.0.0`)

**Features**:

- Builds and packages only the Core library
- Publishes to NuGet as a reusable library (using OIDC trusted publishing)
- Creates GitHub release with library-focused documentation
- Business logic layer for Excel automation
- Shared by MCP Server and CLI

**Release Artifacts**:

- NuGet package: `Sbroenne.ExcelMcp.Core` on NuGet.org
- No binary ZIP (library only, consumed via NuGet)
- Includes dependency on ComInterop

**Publishing Method**:
- Uses OIDC (OpenID Connect) trusted publishing for secure NuGet authentication
- No API keys stored in secrets - authentication via GitHub identity
- Publishes to NuGet.org within the same workflow that creates the GitHub release

**Use Cases**:

- High-level Excel automation commands
- Power Query, VBA, Data Model, worksheets, ranges, tables operations
- Building custom Excel automation tools
- Integrating Excel operations into .NET applications

### 3. MCP Server Releases (`mcp-v*` tags)

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

### 4. CLI Releases (`cli-v*` tags)

**Workflow**: `.github/workflows/release-cli.yml`
**Trigger**: Tags starting with `cli-v` (e.g., `cli-v1.0.0`)

**Features**:

- Builds and packages only the CLI tool
- Publishes to NuGet as a .NET global tool (using OIDC trusted publishing)
- Creates standalone CLI distribution
- Creates GitHub release with CLI-focused documentation
- Focused on direct automation workflows

**Release Artifacts**:

- NuGet package: `Sbroenne.ExcelMcp.CLI` on NuGet.org (as .NET global tool)
- `ExcelMcp-CLI-{version}-windows.zip` - Complete CLI package
- Includes all 89+ commands documentation
- Quick start guide for CLI usage

**Publishing Method**:
- Uses OIDC (OpenID Connect) trusted publishing for secure NuGet authentication
- No API keys stored in secrets - authentication via GitHub identity
- Publishes to NuGet.org within the same workflow that creates the GitHub release
- Package configured as .NET global tool (PackAsTool=true)

**Use Cases**:

- Direct Excel automation scripts
- CI/CD pipeline integration
- Development workflows and testing
- Command-line Excel operations

### 5. VS Code Extension Releases (`vscode-v*` tags)

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

- **ComInterop Library**: Independent version numbers (e.g., cominterop-v1.0.0)
- **Core Library**: Independent version numbers (e.g., core-v1.0.0)
- **MCP Server**: Independent version numbers (e.g., mcp-v1.2.0)
- **CLI**: Independent version numbers (e.g., cli-v2.1.0)
- **VS Code Extension**: Independent version numbers (e.g., vscode-v1.0.5)

### Development Strategy

- **ComInterop**: Focus on COM interop reliability, session management, retry logic
- **Core**: Focus on Excel automation commands, business logic, data operations
- **MCP Server**: Focus on AI integration features, conversational interfaces
- **CLI**: Focus on automation efficiency, command completeness, CI/CD integration
- **VS Code Extension**: Focus on developer experience, VS Code integration, MCP management

## Release Process Examples

### Releasing ComInterop Library Only

```bash
# Create and push ComInterop library release tag
git tag cominterop-v1.0.0
git push origin cominterop-v1.0.0

# This triggers release-cominterop.yml which:
# - Builds ComInterop library
# - Publishes to NuGet using OIDC trusted publishing
# - Creates GitHub release with library-focused docs
```

### Releasing Core Library Only

```bash
# Create and push Core library release tag
git tag core-v1.0.0
git push origin core-v1.0.0

# This triggers release-core.yml which:
# - Builds Core library
# - Publishes to NuGet using OIDC trusted publishing
# - Creates GitHub release with library-focused docs
```

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
# - Publishes to NuGet using OIDC trusted publishing (as .NET global tool)
# - Creates binary distribution
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
- **src/ExcelMcp.ComInterop/README.md**: ComInterop library (low-level COM automation)
- **src/ExcelMcp.Core/README.md**: Core library (high-level Excel commands)
- **docs/CLI.md**: CLI focused (direct automation)
- **vscode-extension/README.md**: VS Code Extension focused (developer experience)
- **Release Notes**: Tailored to the specific component being released

### Cross-References

- Each tool's documentation references the others
- Clear navigation between MCP, CLI, and VS Code Extension docs
- Unified project branding while maintaining component clarity

## Benefits of This Approach

1. **Targeted Releases**: Users can get updates for just the component they use
2. **Independent Development**: Libraries, MCP Server, CLI, and VS Code Extension can evolve at different paces
3. **Reusable Components**: ComInterop and Core can be used in other projects
4. **Focused Documentation**: Release notes and docs match user intent
5. **Reduced Package Size**: Users download only what they need
6. **Clear Separation**: Libraries for building, MCP for AI workflows, CLI for automation, VS Code for IDE
7. **Flexibility**: Each component released independently as needed
8. **NuGet Ecosystem**: Libraries and tools available via NuGet package manager

## Tag Patterns

- `cominterop-v*`: ComInterop library only (NuGet package)
- `core-v*`: Core library only (NuGet package)
- `mcp-v*`: MCP Server only (NuGet .NET tool + binary)
- `cli-v*`: CLI only (NuGet .NET tool + binary)
- `vscode-v*`: VS Code Extension only (VSIX)

This approach provides maximum flexibility while maintaining the integrated ExcelMcp ecosystem.
