# MCPB Build and Packaging Guide

This guide explains how maintainers build the Excel MCP Server bundle for
Claude Desktop. End-user installation instructions are in
[README.md](README.md).

## Directory Contents

```text
mcpb/
|-- Build-McpBundle.ps1   # Packaging script
|-- manifest.json         # MCPB manifest
|-- icon-512.png          # Package icon
|-- README.md             # End-user documentation included in the bundle
|-- BUILD.md              # This maintainer guide
`-- artifacts/            # Generated output (gitignored)
```

## Prerequisites

- .NET 10 SDK
- Windows x64 or Apple Silicon macOS to run matching executable verification

The script can cross-compile either target, but it skips executable launch
verification when the build host cannot run that target.

## Build the Bundle

Run the script from the `mcpb` directory:

```powershell
.\Build-McpBundle.ps1 -RuntimeIdentifier win-x64
.\Build-McpBundle.ps1 -RuntimeIdentifier osx-arm64
```

The outputs are `artifacts\excel-mcp-{version}-windows.mcpb` and
`artifacts\excel-mcp-{version}-macos-arm64.mcpb`. The version comes from
`Directory.Build.props` unless you pass it explicitly.

```powershell
# Use an explicit version
.\Build-McpBundle.ps1 -Version "1.2.3" -RuntimeIdentifier osx-arm64

# Write artifacts to another directory
.\Build-McpBundle.ps1 -OutputDir ".\dist"
```

## Package Contents

An `.mcpb` file is a ZIP-compatible archive with this layout:

```text
excel-mcp-{version}-{platform}.mcpb
|-- manifest.json
|-- icon-512.png
|-- README.md
|-- LICENSE
|-- CHANGELOG.md
`-- server/
    `-- excel-mcp-server[.exe]
```

`Build-McpBundle.ps1` publishes exactly one self-contained native executable,
stamps the matching `win32` or `darwin` compatibility metadata, and preserves
the Unix executable mode in the macOS archive.

## Manifest and Tool Metadata

`manifest.json` follows MCPB manifest version 0.3 and declares the packaged
binary entry point:

```json
{
  "manifest_version": "0.3",
  "server": {
    "type": "binary",
    "entry_point": "server/excel-mcp-server",
    "mcp_config": {
      "command": "${__dirname}/server/excel-mcp-server",
      "args": [],
      "env": {}
    }
  }
}
```

The build stamps the package version, display name, description, and one
platform into a staged copy of the manifest. Each bundle contains only its
matching executable.

The MCP Server generates its 31 tool schemas from the Core contracts and manual
MCP tool definitions. Destructive metadata is set per tool: most tools can
modify workbooks, while tools such as `screenshot` and `window` do not modify
workbook content.

## Release Workflow

The unified release workflow builds and publishes all three MCPB artifacts with the
MCP Server, CLI, VS Code extension, and NuGet packages. Do not edit the manifest
or upload a differently named ZIP by hand.

See [Release Strategy](../docs/RELEASE-STRATEGY.md) for the release process. To
rebuild locally before a release:

```powershell
.\Build-McpBundle.ps1 -RuntimeIdentifier win-x64
.\Build-McpBundle.ps1 -RuntimeIdentifier osx-arm64
```

## Verify the Archive

PowerShell's `Expand-Archive` expects a `.zip` extension. Copy the generated
bundle before inspecting it:

```powershell
$bundle = Get-ChildItem .\artifacts\excel-mcp-*.mcpb |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1
$zip = [IO.Path]::ChangeExtension($bundle.FullName, ".zip")
Copy-Item $bundle.FullName $zip
Expand-Archive $zip -DestinationPath .\test-extract
Get-ChildItem .\test-extract -Recurse
Remove-Item $zip
Remove-Item .\test-extract -Recurse
```

The packaging script also prints every archive entry after a successful build.

## Technical Notes

### Why Self-Contained?

- Users do not need the .NET runtime or SDK.
- The package uses the same tested runtime as the standalone release.
- A single executable avoids local dependency version conflicts.

### Why No Trimming?

Excel COM interop relies on runtime type activation and reflection. Trimming can
remove required interop metadata, so the package sets `PublishTrimmed=false`.

### Why separate bundles?

- PE and Mach-O are different native executable formats.
- Claude Desktop receives only the runtime for the current platform.
- The macOS archive must retain its Unix executable permission.

## Submission References

- [MCPB submission guide](https://support.claude.com/en/articles/12922832-local-mcp-server-submission-guide)
- [Claude Desktop documentation](https://support.claude.com/)
- [MCP documentation](https://modelcontextprotocol.io/)
