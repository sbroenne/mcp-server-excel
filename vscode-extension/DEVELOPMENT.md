# VS Code Extension Development

The Excel MCP Server extension bundles the Windows x64 MCP executable and one
Agent Skill. It ships separate Windows x64 and Windows ARM64 VSIX packages;
ARM64 Windows runs the same executable through x64 emulation. Users do not
need a separate .NET runtime, Node.js runtime, or CLI installation.

## Project structure

```text
vscode-extension/
├── src/extension.ts          # Extension activation and MCP registration
├── out/                      # Compiled JavaScript
├── bin/                      # Self-contained MCP Server built for packaging
├── skills/excel-mcp/         # Build copy of the canonical Agent Skill
├── scripts/                  # Build-time manifest validation
├── tests/                    # Vitest registration and setup regressions
├── vitest.config.mts         # Node tests with a test-only VS Code API replacement
├── tsconfig.test.json        # Separate test-source type checking
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
shared package command copies the prepared MCP skill from
`artifacts/generated-skills/` into an isolated extension staging directory.

After a Release solution build, `npm run package` uses
`scripts/Build-ReleasePackages.ps1 -Components Extension`. It publishes the MCP
runtime once, generates complete skills, installs locked extension dependencies,
and creates and inspects the VSIX under `artifacts/packages/`. It does not clean
or overwrite this source directory. For an unpackaged debug session, open the
prepared `extension` directory reported by that command.

Do not edit `vscode-extension/CHANGELOG.md` directly. The build copies the
generated root `CHANGELOG.md` into the extension package.

## Contributions

### MCP Server

The manifest declares the provider ID and `extensionKind: ["ui"]`. The
extension registers that exact ID through VS Code's API and includes its
release-stamped version so upgrades refresh tool discovery:

```typescript
vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
  provideMcpServerDefinitions: async () => [
    new vscode.McpStdioServerDefinition(
      'excel-mcp',
      path.join(context.extensionPath, 'bin', 'Sbroenne.ExcelMcp.McpServer.exe'),
      [],
      {},
      context.extension.packageJSON.version
    )
  ]
});
```

Discovery does not probe the installation or prompt the user.
`resolveMcpServerDefinition` checks Windows, executable access, and Excel COM
registration just before launch. The registration check uses noninteractive
Windows PowerShell with a ten-second timeout and cancellation; it does not
create Excel or open a workbook.

The first-run Getting Started action opens the published user guides, not
installation instructions for an extension that is already installed.

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
- Node.js 22.12+ and npm, compatible with the locked Vitest and vsce versions
- Microsoft Excel for end-to-end MCP testing

Install locked dependencies:

```powershell
Set-Location vscode-extension
npm ci
```

## Build and validation

Run the fast checks while editing:

```powershell
npm test
npm run typecheck:tests
npm run lint
```

Vitest runs once with `vitest run`. It exercises the extension's actual
registration and setup code with test-only VS Code API and prerequisite
replacements. It does not establish Excel COM behavior. Tests, mocks, test
configuration, caches, and reports do not ship.

`npm run compile` validates Marketplace assets and feature metadata before
compiling production TypeScript. It needs the generated skill inputs present
in the prepared package stage; the source directory deliberately does not
contain those generated copies.

Build the complete release-shaped VSIX:

```powershell
npm run package
```

Packaging performs these steps automatically:

1. Publishes the MCP Server as a self-contained Windows x64 executable.
2. Copies and stamps the canonical Agent Skill.
3. Copies the generated root changelog.
4. Validates feature and Marketplace metadata and compiles TypeScript.
5. Runs lint, test-source type checking, and the focused Vitest suite.
6. Runs `vsce package --target win32-x64` and `--target win32-arm64`.
7. Inspects both VSIX targets, versions, all skill files, executable, compiled
   source, and exclusions for development files and the CLI.

To prepare and inspect the bundled executable through the shared package path,
run these commands from the repository root:

```powershell
.\scripts\Build-AgentSkills.ps1 -GenerateOnly
.\scripts\Build-ReleasePackages.ps1 -Components Extension -SkillsDirectory artifacts\generated-skills -OutputDirectory artifacts\extension-check
.\artifacts\extension-check\runtimes\Mcp\Sbroenne.ExcelMcp.McpServer.exe --version
```

## Local testing

### Extension Development Host

1. Run `npm run package` from the source extension folder.
2. Open the prepared `extension` directory reported by packaging in VS Code.
3. Install its locked development dependencies with `npm ci` if the debug
   configuration needs a compile step, then press `F5`.
4. Confirm the MCP server appears in VS Code's MCP management UI.
5. Open the extension's Features tab and verify:
   - MCP Servers shows `excel-mcp` and `Excel MCP Server`.
   - Chat Skills shows the skill name, description, and path.

### Packaged VSIX

1. Run `npm run package`.
2. Use **Extensions: Install from VSIX...** in VS Code.
3. Select `excel-mcp-<version>.vsix` for Windows x64 or
   `excel-mcp-<version>-win32-arm64.vsix` for native ARM64 VS Code.
4. Reload VS Code and verify the MCP server and Agent Skill.

The VSIX is approximately 65 MB. Most of its size is the compressed,
self-contained MCP Server; the unpacked executable is approximately 150 MB.

For an isolated check, use separate `--user-data-dir` and `--extensions-dir`
directories rather than overwrite your everyday installation. Check activation,
first-run help, the Features tab, and **MCP: List Servers > excel-mcp > Show
Output**. The **ExcelMcp** output channel contains setup diagnostics, not server
operation logs.

Use a temporary workbook for a create/read/close smoke with real Excel. Run
Excel-dependent checks sequentially. In a remote workspace, verify the server
runs on the Windows desktop and requires locally accessible workbook paths.

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

Run the shared package commands above, then verify that
`artifacts/extension-check/runtimes/Mcp/Sbroenne.ExcelMcp.McpServer.exe --version`
succeeds. The manifest provider
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
