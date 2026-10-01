# Claude MCPB Submission Guide

## Purpose
Submit Excel MCP Server to Anthropic’s Claude Directory as an MCPB bundle for one-click installation in Claude Desktop.

## Prerequisites
- MCPB bundle built and validated
- 512×512 PNG icon available
- Privacy page published
- Separate authorization to publish or submit; local package checks are not publication

## Required Assets
- MCPB bundles: Windows npx metadata and Apple Silicon macOS native GitHub Actions artifacts
- MCPB manifest: mcpb/manifest.json
- Icon: mcpb/icon-512.png
- Privacy page: https://excelmcpserver.dev/privacy/

## Build Steps
1. Run the release workflow to produce both MCPB artifacts.
2. Download both MCPB artifacts from the workflow run.
3. Verify the Windows artifact contains only metadata and launches
   `@sbroenne/mcp-server-excel@latest`; verify the Mac artifact contains only
   its signed native executable and helper.

## Tool Annotation Requirement
The C# MCP SDK maps tool hints from [McpServerTool] attribute properties:
- Destructive = true → annotations.destructiveHint = true
- ReadOnly, Idempotent, OpenWorld map similarly

Nearly all tools set Destructive = true, since Excel automation modifies live workbook state; read-only tools (e.g. `screenshot`) set Destructive = false.

## Submission Form Checklist
Fill the Claude Directory submission form with:
- Server name: Excel MCP Server
- MCPB files: downloaded Windows npx and Apple Silicon macOS native artifacts
- Website: https://excelmcpserver.dev/
- Privacy policy: https://excelmcpserver.dev/privacy/
- Support or repo link: https://github.com/sbroenne/mcp-server-excel
- Icon: mcpb/icon-512.png
- Platform notes: Windows x64 uses the complete COM backend; Apple Silicon
  macOS is an **experimental beta**, not full parity. Include the
  [unsupported-feature reference](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta);
  optional handlers do not constitute supported features

## Post-Submission
- Record submission timestamp and form confirmation URL in the GitHub issue
- If requested, attach the MCPB bundle and icon to the issue
