# ExcelMcp Release Strategy

This document outlines the separate build and release processes for the ExcelMcp MCP Server and CLI components.

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

### 3. Combined Releases (`v*` tags)

**Workflow**: `.github/workflows/release.yml`
**Trigger**: Tags starting with `v` (e.g., `v3.0.0`)

**Features**:

- Builds both MCP Server and CLI
- Creates combined distribution package
- Publishes MCP Server to NuGet (using OIDC trusted publishing)
- Comprehensive release with both tools
- Maintains backward compatibility

**Release Artifacts**:

- `ExcelMcp-{version}-windows.zip` - Combined package
- NuGet package: `Sbroenne.ExcelMcp.McpServer` on NuGet.org
- Contains both CLI and MCP Server binaries
- Unified documentation and installation guide

**Use Cases**:

- Users who need both tools
- Complete ExcelMcp installation
- Comprehensive Excel development environment

## Version Management

### Independent Versioning

- **MCP Server**: Can have independent version numbers (e.g., mcp-v1.2.0)
- **CLI**: Can have independent version numbers (e.g., cli-v2.1.0)
- **Combined**: Major releases combining both (e.g., v3.0.0)

### Development Strategy

- **MCP Server**: Focus on AI integration features, conversational interfaces
- **CLI**: Focus on automation efficiency, command completeness, CI/CD integration
- **Combined**: Major milestones, breaking changes, coordinated releases

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

### Combined Release

```bash
# Create and push combined release tag
git tag v3.1.0
git push origin v3.1.0

# This triggers release.yml which:
# - Builds both MCP server and CLI
# - Publishes MCP Server to NuGet using OIDC trusted publishing
# - Creates combined distribution package
# - Creates GitHub release with comprehensive documentation
```

## Documentation Strategy

### Separate Focus Areas

- **Main README.md**: MCP Server focused (AI assistant integration)
- **docs/CLI.md**: CLI focused (direct automation)
- **Release Notes**: Tailored to the specific component being released

### Cross-References

- Each tool's documentation references the other
- Clear navigation between MCP and CLI docs
- Unified project branding while maintaining component clarity

## Benefits of This Approach

1. **Targeted Releases**: Users can get updates for just the tool they use
2. **Independent Development**: MCP and CLI can evolve at different paces
3. **Focused Documentation**: Release notes and docs match user intent
4. **Reduced Package Size**: Users download only what they need
5. **Clear Separation**: MCP for AI workflows, CLI for direct automation
6. **Flexibility**: Combined releases still available for comprehensive updates

## Migration from Single Release

### Existing Tags

- Previous `v*` tags remain as combined releases
- No breaking changes to existing release structure

### New Tag Patterns

- `mcp-v*`: MCP Server only
- `cli-v*`: CLI only  
- `v*`: Combined (maintains compatibility)

This approach provides maximum flexibility while maintaining the integrated ExcelMcp ecosystem.
