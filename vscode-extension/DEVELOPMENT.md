# VS Code Extension Development

The Excel MCP Server extension ships platform-targeted VSIX packages for
Windows x64 and Apple Silicon macOS. Each package bundles the matching MCP
executable and one Agent Skill. Users do not need a separate .NET runtime or
CLI installation.

Mac support is an experimental beta subset, not full Windows parity; keep
package metadata and examples aligned with
[the support reference](../specs/MACOS-SUPPORT.md).

## Project structure

```text
vscode-extension/
├── src/extension.ts          # Extension activation and MCP registration
├── out/                      # Compiled JavaScript
├── bin/                      # Self-contained MCP Server built for packaging
├── skills/excel-mcp-report-formatting/ # Build copy of the formatting skill
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

Do not edit files under `vscode-extension/skills/excel-mcp-report-formatting/` directly. The
shared package command copies the prepared MCP skill from
`artifacts/generated-skills/` into an isolated extension staging directory.

On Windows, after a Release solution build, `npm run package` uses
`scripts/Build-ReleasePackages.ps1 -Components Extension`. It publishes the MCP
runtime once, generates complete skills, installs locked extension dependencies,
and creates and inspects the VSIX under `artifacts/packages/`. It does not clean
or overwrite this source directory. For an unpackaged debug session, open the
prepared `extension` directory reported by that command.

Do not edit `vscode-extension/CHANGELOG.md` directly. The build copies the
generated root `CHANGELOG.md` into the extension package.

## Contributions

### MCP Server

The manifest declares the provider ID and the extension registers that exact
ID through VS Code's API:

The platform resolver selects `win32-x64/Sbroenne.ExcelMcp.McpServer.exe` or
`darwin-arm64/Sbroenne.ExcelMcp.McpServer`; unsupported hosts fail explicitly:

```typescript
vscode.lm.registerMcpServerDefinitionProvider('excel-mcp', {
  provideMcpServerDefinitions: async () => {
    const runtime = resolveBundledRuntime(process.platform, process.arch);
    return [
      new vscode.McpStdioServerDefinition(
        'excel-mcp',
        path.join(context.extensionPath, 'bin', runtime.directory, runtime.executable),
        [],
        {}
      )
    ];
  }
});
```

### Agent Skill

The `chatSkills` contribution registers the packaged skill:

```json
"chatSkills": [
  {
    "name": "excel-mcp-report-formatting",
    "description": "Optional presentation conventions for requested Excel reports.",
    "path": "./skills/excel-mcp-report-formatting/SKILL.md"
  }
]
```

VS Code's Extension Features table reads `name` and `description` from the
manifest. Runtime skill discovery reads the matching frontmatter from
`SKILL.md`; the build validates that the names agree.

## Prerequisites

- Windows x64 or Apple Silicon macOS
- The .NET SDK pinned by the repository `global.json`
- Node.js and npm
- PowerShell 7 for the packaging scripts
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

Vitest runs once with `vitest run`. It exercises the extension's actual
registration and setup code with test-only VS Code API and prerequisite
replacements. It does not establish Excel COM behavior. Tests, mocks, test
configuration, caches, reports, and developer-only `AGENTS.md`/`CLAUDE.md`
instructions do not ship, including nested instruction files.

Build the Windows release-shaped VSIX on Windows:

```powershell
npm run package
```

Windows packaging performs these steps automatically:

1. Publishes the self-contained Windows x64 MCP runtime.
2. Copies and stamps the canonical Agent Skill.
3. Copies the generated root changelog.
4. Validates feature and Marketplace metadata and compiles TypeScript.
5. Runs lint, test-source type checking, and the focused Vitest suite.
6. Copies the matching server before each `vsce package --target win32-x64` or
   `--target win32-arm64` invocation.
7. Inspects both VSIX targets, versions, all skill files, the actual bundled
   executable's CPU architecture, compiled source, and exclusions for development
   files, developer instructions, and the CLI. A mislabeled server or leaked
   instruction file fails packaging.

Apple Silicon release packaging runs separately from the repository root:

```powershell
./scripts/Build-MacReleasePackages.ps1 -Version <version> -OutputDirectory artifacts/release-macos
```

That path builds, signs, launches, and inspects the ARM64 runtime and produces
the `darwin-arm64` VSIX. Intel macOS is unsupported and has no VSIX target.

On Windows, to prepare and inspect the bundled executable through the shared package path,
run these commands from the repository root:

```powershell
.\scripts\Build-AgentSkills.ps1 -GenerateOnly
.\scripts\Build-ReleasePackages.ps1 -Components Extension -SkillsDirectory artifacts\generated-skills -OutputDirectory artifacts\extension-check
.\artifacts\extension-check\runtimes\Mcp\Sbroenne.ExcelMcp.McpServer.exe --version
```

## Local testing

### Extension Development Host

1. Open the prepared extension staging directory for your platform in VS Code
   (it contains the generated skill and bundled runtime).
2. Run `npm run compile`.
3. Press `F5` to open an Extension Development Host.
4. Confirm the MCP server appears in VS Code's MCP management UI.
5. Open the extension's Features tab and verify:
   - MCP Servers shows `excel-mcp` and `Excel MCP Server`.
   - Chat Skills shows the skill name, description, and path.

### Packaged VSIX

1. Run `npm run package` on Windows or the Mac packaging command above on Apple Silicon.
2. Use **Extensions: Install from VSIX...** in VS Code.
3. Select `excelmcp-<version>-win32-x64.vsix` or
   `excelmcp-<version>-darwin-arm64.vsix`.
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

On Windows, run the shared package commands above, then verify that
`artifacts/extension-check/runtimes/Mcp/Sbroenne.ExcelMcp.McpServer.exe --version`
succeeds. The manifest provider
ID and the ID passed to `registerMcpServerDefinitionProvider` must both be
`excel-mcp`.

On Mac, use the native packaging path and check
`<output>/runtimes/Mcp/Sbroenne.ExcelMcp.McpServer --version`. Do not run the
Windows package path or expect an `.exe`.

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
