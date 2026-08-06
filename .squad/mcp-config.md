# MCP Integration — Configuration and Samples

MCP (Model Context Protocol) servers extend Squad with tools for external services — Trello, Aspire dashboards, Notion, and more. The user configures MCP servers in their environment; Squad discovers and uses them.

## Config File Locations

Users configure MCP servers at these locations (checked in priority order):
1. **Repository-level:** `.copilot/mcp-config.json` (team-shared, committed to repo)
2. **Workspace-level:** `.vscode/mcp.json` (VS Code workspaces)
3. **User-level:** `~/.copilot/mcp-config.json` (personal)
4. **CLI override:** `--additional-mcp-config` flag (session-specific)

## Sample Config — Trello

```json
{
  "mcpServers": {
    "trello": {
      "command": "npx",
      "args": ["-y", "@trello/mcp-server"],
      "env": {
        "TRELLO_API_KEY": "${TRELLO_API_KEY}",
        "TRELLO_TOKEN": "${TRELLO_TOKEN}"
      }
    }
  }
}
```

## Sample Config — GitHub

```json
{
  "mcpServers": {
    "github": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-github"],
      "env": {
        "GITHUB_TOKEN": "${GITHUB_TOKEN}"
      }
    }
  }
}
```

## Sample Config — Aspire

```json
{
  "mcpServers": {
    "aspire": {
      "command": "npx",
      "args": ["-y", "@aspire/mcp-server"],
      "env": {
        "ASPIRE_DASHBOARD_URL": "${ASPIRE_DASHBOARD_URL}"
      }
    }
  }
}
```

## Authentication Notes

- **GitHub MCP requires a separate token** from the `gh` CLI auth. Generate at https://github.com/settings/tokens
- **Trello requires API key + token** from https://trello.com/power-ups/admin
- **Aspire uses the dashboard URL** — typically `http://localhost:18888` during local dev

Auth is a real blocker for some MCP servers. Users need separate tokens for GitHub MCP, Trello MCP, etc. This is a documentation problem, not a code problem.
