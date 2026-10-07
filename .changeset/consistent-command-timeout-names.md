---
"excelmcp": patch
---

**Consistent command and timeout names** (#1066): `excelcli batch` now accepts the same command group names as the CLI (`calculationmode`, `datamodelrelationship`, `worksheetstyle`). Timeouts are named `--timeout-seconds` in the CLI, `timeoutSeconds` in batch JSON, and `timeout_seconds` in MCP. MCP `file` and `file_read` now take `file_path` instead of `path`. Mistyped commands or arguments now return a list of the valid choices. These are breaking changes; see [Breaking Changes](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/BREAKING-CHANGES.md).
