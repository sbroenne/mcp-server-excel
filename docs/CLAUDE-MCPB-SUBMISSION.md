# Claude MCPB Submission Guide

## Purpose
Submit Excel MCP Server to Anthropic’s Claude Directory as an MCPB bundle for one-click installation in Claude Desktop.

## Prerequisites
- MCPB bundle built and validated
- 512×512 PNG icon available
- Privacy page published

## Required Assets
- MCPB bundle: GitHub Actions release workflow artifact (.mcpb)
- MCPB manifest: mcpb/manifest.json
- Icon: mcpb/icon-512.png
- Privacy page: https://excelmcpserver.dev/privacy/

## Build Steps
1. Run the release workflow to produce the MCPB artifact.
2. Download the MCPB artifact from the workflow run.
3. Verify the artifact is the intended .mcpb bundle for submission.

## Tool Annotation Requirement
The C# MCP SDK maps tool hints from [McpServerTool] attribute properties:
- Destructive = true → annotations.destructiveHint = true
- ReadOnly, Idempotent, OpenWorld map similarly

All 22 tools in this repository set Destructive = true.

## Submission Form Checklist
Fill the Claude Directory submission form with:
- Server name: Excel MCP Server
- MCPB file: downloaded workflow artifact (.mcpb)
- Website: https://excelmcpserver.dev/
- Privacy policy: https://excelmcpserver.dev/privacy/
- Support or repo link: https://github.com/sbroenne/mcp-server-excel
- Icon: mcpb/icon-512.png
- Platform notes: Windows-only (Excel COM), x64 self-contained build

## Post-Submission
- Record submission timestamp and form confirmation URL in the GitHub issue
- If requested, attach the MCPB bundle and icon to the issue
