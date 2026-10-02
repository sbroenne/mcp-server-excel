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

- PowerShell 7
- Node.js/npm when verifying the server launch (not needed to build the archive)

Packaging copies metadata only. It does not publish .NET executables, install
npm dependencies, or download the server.

## Build the Bundle

Run the script from the `mcpb` directory:

```powershell
.\Build-McpBundle.ps1
```

The default output is `artifacts\excel-mcp-{version}.mcpb`. The version comes
from `Directory.Build.props` unless you pass it explicitly.

```powershell
# Use an explicit version
.\Build-McpBundle.ps1 -Version "1.2.3"

# Write artifacts to another directory
.\Build-McpBundle.ps1 -OutputDir ".\dist"
```

## Package Contents

An `.mcpb` file is a ZIP-compatible archive with this layout:

```text
excel-mcp-{version}.mcpb
|-- manifest.json
|-- icon-512.png
|-- README.md
|-- LICENSE
`-- CHANGELOG.md
```

`Build-McpBundle.ps1` stamps a staged manifest, copies the package metadata,
checks the archive entries, and installs the completed output. Failed builds
preserve existing packages; staging cleanup is restricted to the owned
temporary directory.

## Manifest and Tool Metadata

`manifest.json` follows MCPB manifest version 0.3. The entry point identifies
the npm package, while `mcp_config` specifies the actual command:

```json
{
  "manifest_version": "0.3",
  "server": {
    "type": "node",
    "entry_point": "@sbroenne/mcp-server-excel",
    "mcp_config": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"],
      "env": {}
    }
  }
}
```

The build stamps the package version into a staged copy of the manifest. Do not
add a release download URL or an `install.win32` block. npx obtains the server
from npm at launch, subject to normal npm configuration and caching.

The MCP Server generates its 31 tool schemas from the Core contracts and manual
MCP tool definitions. Destructive metadata is set per tool: most tools can
modify workbooks, while tools such as `screenshot` and `window` do not modify
workbook content.

## Release Workflow

The unified release workflow builds and publishes the MCPB artifact with the
MCP Server, CLI, VS Code extension, and NuGet packages. Do not edit the manifest
or upload a differently named ZIP by hand.

See [Release Strategy](../docs/RELEASE-STRATEGY.md) for the release process. To
rebuild locally before a release:

```powershell
.\Build-McpBundle.ps1
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

Verify the actual configured command on Windows as well:

```powershell
npx -y @sbroenne/mcp-server-excel@latest --version
```

This checks startup, not full Claude integration. Install the bundle in Claude
Desktop and confirm initialization, tool discovery, and a create/save/close
workbook operation before claiming host integration is verified.

## Technical Notes

### Why Direct npx?

- No custom launcher or bundled npm dependencies.
- The installed bundle does not pin the server to its stamped release version.
- The npm server contains a self-contained .NET executable.
- Node.js/npm must be available on PATH; do not assume Claude bundles npx.

### Updates and Network Access

The bundle is not offline/self-contained. First launch downloads the server;
later launches use normal npm resolution and caching. The `latest` tag can
resolve to a different version than the bundle's metadata version. A running
server is never hot-swapped. Existing binary MCPB users must install this
configuration bundle once, and future bundle changes still require replacement.

### Architecture and Directory Submission

- Excel COM automation requires Windows.
- npm selects the x64 or ARM64 runtime matching the Node.js process architecture.
- x64 Node.js on ARM64 Windows uses the x64 package through emulation.
- Directory acceptance of a fetch-on-launch bundle is not verified. Disclose
  its npm and network requirements; do not claim it bundles all dependencies.

## Submission References

- [MCPB submission guide](https://support.claude.com/en/articles/12922832-local-mcp-server-submission-guide)
- [Claude Desktop documentation](https://support.claude.com/)
- [MCP documentation](https://modelcontextprotocol.io/)
