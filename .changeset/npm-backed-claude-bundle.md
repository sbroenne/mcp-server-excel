---
"excelmcp": major
---

**Latest-server launches from Claude Desktop**: The MCPB now configures `npx -y @sbroenne/mcp-server-excel@latest` directly instead of bundling a fixed executable. Node.js with npm/npx must be available on PATH; no separate .NET installation is needed. Existing binary-bundle users must install the new bundle once. Server versions are resolved on launch using normal npm caching; running workbook sessions are not automatically restarted.

Installation guides and ready-to-use MCP configurations now consistently use `@latest` and explain requirements, updates, removal, and safe CLI service restarts.
