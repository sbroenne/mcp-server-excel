# VS Code Extension Development

The Excel MCP Server extension bundles the Windows x64 MCP executable and one
Agent Skill. Users do not need a separate .NET runtime or CLI installation.

## Project structure

```text
vscode-extension/
├── src/extension.ts          # Extension activation and MCP registration
├── out/                      # Compiled JavaScript
├── bin/                      # Self-contained MCP Server built for packaging
├── skills/excel-mcp/         # Build copy of the canonical Agent Skill
├── scripts/                  # Build-time manifest validation
├── package.json              # Extension manifest and npm scripts
├── package-lock.json         # Locked development dependencies
├── tsconfig.json             # TypeScript compiler settings
├── .vscodeignore             # Files excluded from the VSIX
├── README.md                 # Marketplace details page
├── CHANGELOG.md              # Build copy of the repository changelog
├── LICENSE                   # Packaged MIT license
└── icon.png                  # Marketplace icon
```

Do not edit files under `vscode-extension/skills/excel-mcp/` directly. The
`copy:skills` script replaces that directory from the canonical
`skills/excel-mcp/` source and stamps `VERSION` from `package.json`.

Do not edit `vscode-extension/CHANGELOG.md` directly. The build copies the
generated root `CHANGELOG.md` into the extension package.

## Contributions

### MCP Server

The manifest declares the provider ID and the extension registers that exact
ID through VS Code's API:

```typescript
vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
  provideMcpServerDefinitions: async () => [
    new vscode.McpStdioServerDefinition(
      'excel-mcp',
      path.join(context.extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe'),
      [],
      {}
    )
  ]
});
```

### Agent Skill

The `chatSkills` contribution registers the packaged skill:

```json
"chatSkills": [
  {
    "name": "excel-mcp",
    "description": "Excel MCP Server skill for Windows workbook automation.",
    "path": "./skills/excel-mcp/SKILL.md"
  }
]
```

VS Code's Extension Features table reads `name` and `description` from the
manifest. Runtime skill discovery reads the matching frontmatter from
`SKILL.md`; the build validates that the names agree.

## Prerequisites

- Windows
- The .NET SDK pinned by the repository `global.json`
- Node.js and npm
- Microsoft Excel for end-to-end MCP testing

Install locked dependencies:

```powershell
Set-Location vscode-extension
npm ci
```

## Build and validation

Run the fast checks while editing:

```powershell
npm run compile
npm run lint
```

`compile` first validates Marketplace assets and feature contribution metadata,
then compiles TypeScript.

Build the complete release-shaped VSIX:

```powershell
npm run package
```

Packaging performs these steps automatically:

1. Publishes the MCP Server as a self-contained Windows x64 executable.
2. Copies and stamps the canonical Agent Skill.
3. Copies the generated root changelog.
4. Validates feature and Marketplace metadata.
5. Compiles TypeScript and runs `vsce package`.

To build only the bundled executable:

```powershell
npm run build:mcp-server
.\bin\Sbroenne.ExcelMcp.McpServer.exe --version
```

## Local testing

### Extension Development Host

1. Open the `vscode-extension` folder in VS Code.
2. Run `npm run compile`.
3. Press `F5` to open an Extension Development Host.
4. Confirm the MCP server appears in VS Code's MCP management UI.
5. Open the extension's Features tab and verify:
   - MCP Servers shows `excel-mcp` and `Excel MCP Server`.
   - Chat Skills shows the skill name, description, and path.

### Packaged VSIX

1. Run `npm run package`.
2. Use **Extensions: Install from VSIX...** in VS Code.
3. Select the generated `excel-mcp-<version>.vsix`.
4. Reload VS Code and verify the MCP server and Agent Skill.

The VSIX is approximately 65 MB. Most of its size is the compressed,
self-contained MCP Server; the unpacked executable is approximately 150 MB.

## Release workflow

For every user-visible extension change:

1. Add a patch changeset from the repository root with `npx changeset`.
2. Do not manually edit package versions or either changelog copy.
3. Open a pull request and let CI validate the package.
4. Use the unified **Release All Components** workflow after merge.

The release workflow calculates one version for all deliverables, compiles the
changesets into the root changelog, packages the extension, publishes it to the
VS Code Marketplace, publishes the .NET packages, and creates the GitHub
release.

## Troubleshooting

### Missing Node modules

Run `npm ci` from `vscode-extension`.

### MCP Server is missing

Run `npm run build:mcp-server`, then verify that
`bin/Sbroenne.ExcelMcp.McpServer.exe --version` succeeds. The manifest provider
ID and the ID passed to `registerMcpServerDefinitionProvider` must both be
`excel-mcp`.

### Feature metadata validation fails

Keep the provider ID and label populated. Keep every Chat Skill's manifest
name synchronized with the `name` field in its `SKILL.md` frontmatter, and
provide a manifest description for VS Code's Features table.

### Package contents are wrong

Review `.vscodeignore`, then run:

```powershell
npx vsce ls --tree
```

Development sources, local VSIX files, dependencies, and build-only scripts
must not ship in the package.

## References

- [VS Code Extension API](https://code.visualstudio.com/api)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Publishing Extensions](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
