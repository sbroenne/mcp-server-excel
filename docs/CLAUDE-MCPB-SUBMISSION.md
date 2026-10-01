# Claude MCPB Submission Guide

## Purpose
Submit Excel MCP Server to Anthropic’s Claude Directory as an MCPB bundle for one-click installation in Claude Desktop.

## Prerequisites
- MCPB bundle built and validated
- Direct npx launch tested in Claude Desktop on Windows with Node.js/npm on PATH
- 512×512 PNG icon available
- Privacy page published

## Required Assets
- MCPB bundle: GitHub Actions release workflow artifact (.mcpb)
- MCPB manifest: mcpb/manifest.json
- Icon: mcpb/icon-512.png
- Privacy page: https://excelmcpserver.dev/privacy/

## Build Steps
1. Build locally with `.\mcpb\Build-McpBundle.ps1`, or use the authorized release artifact.
2. Inspect the `.mcpb`: it contains metadata and a direct
   `npx -y @sbroenne/mcp-server-excel@latest` configuration, not a fixed executable.
3. Test initialization, tool discovery, and a create/save/close workbook operation
   in Claude Desktop. Do not dispatch a release workflow merely to test packaging.

Directory acceptance of this npm fetch-on-launch design is not verified. The
bundle does not include all runtime dependencies: Node.js/npm must be available
on PATH, and npx needs network access for downloads and update checks. Disclose
this in the submission rather than describing the bundle as self-contained.

## Tool Annotation Requirement
The C# MCP SDK maps tool hints from [McpServerTool] attribute properties:
- Destructive = true → annotations.destructiveHint = true
- ReadOnly, Idempotent, OpenWorld map similarly

Nearly all tools set Destructive = true, since Excel automation modifies live workbook state; read-only tools (e.g. `screenshot`) set Destructive = false.

## Submission Form Checklist
Fill the Claude Directory submission form with:
- Server name: Excel MCP Server
- MCPB file: downloaded workflow artifact (.mcpb)
- Website: https://excelmcpserver.dev/
- Privacy policy: https://excelmcpserver.dev/privacy/
- Support or repo link: https://github.com/sbroenne/mcp-server-excel
- Icon: mcpb/icon-512.png
- Platform notes: Windows-only (Excel COM); npx selects the x64 or ARM64 npm
  runtime matching Node.js. Requires Node.js/npm and network access; no separate .NET.

## Post-Submission
- Record submission timestamp and form confirmation URL in the GitHub issue
- If requested, attach the MCPB bundle and icon to the issue
