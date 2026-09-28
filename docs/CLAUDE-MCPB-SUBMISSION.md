# Claude MCPB Submission Guide

## Purpose
Submit Excel MCP Server to Anthropic’s Claude Directory as an MCPB bundle for one-click installation in Claude Desktop.

## Prerequisites
- MCPB bundle built and validated
- 512×512 PNG icon available
- Privacy page published

## Required Assets
- MCPB bundles: Windows x64 and macOS x64/Arm64 GitHub Actions artifacts
- MCPB manifest: mcpb/manifest.json
- Icon: mcpb/icon-512.png
- Privacy page: https://excelmcpserver.dev/privacy/

## Build Steps
1. Run the release workflow to produce all three MCPB artifacts.
2. Download all three MCPB artifacts from the workflow run.
3. Verify each artifact contains only its platform's native executable.

## Tool Annotation Requirement
The C# MCP SDK maps tool hints from [McpServerTool] attribute properties:
- Destructive = true → annotations.destructiveHint = true
- ReadOnly, Idempotent, OpenWorld map similarly

Nearly all tools set Destructive = true, since Excel automation modifies live workbook state; read-only tools (e.g. `screenshot`) set Destructive = false.

## Submission Form Checklist
Fill the Claude Directory submission form with:
- Server name: Excel MCP Server
- MCPB files: downloaded Windows x64 and macOS x64/Arm64 artifacts
- Website: https://excelmcpserver.dev/
- Privacy policy: https://excelmcpserver.dev/privacy/
- Support or repo link: https://github.com/sbroenne/mcp-server-excel
- Icon: mcpb/icon-512.png
- Platform notes: Windows x64 uses the complete COM backend; macOS x64/Arm64
  uses the documented capability-gated Apple Events backend. Intel packages
  are cross-built and hardware-unverified

## Post-Submission
- Record submission timestamp and form confirmation URL in the GitHub issue
- If requested, attach the MCPB bundle and icon to the issue
