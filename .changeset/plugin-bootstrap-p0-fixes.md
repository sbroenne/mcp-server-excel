---
"excelmcp": patch
---

**Copilot plugins: arguments with quotes now survive, and a cached runtime keeps
working offline.** Two defects made the published `excel-cli` and `excel-mcp`
plugins fail in normal use.

Every documented inline JSON example — `--values '[["Name","Amount"]]'` — was
silently corrupted to `[[Name,Amount]]` when invoked through the plugin's own
wrapper or the generated `excelcli` PATH shim, because Windows PowerShell rebuilds
the command line for native executables and drops embedded double quotes. The
wrapper now builds the command line itself using the standard MSVCRT quoting rules
and hands it to the process verbatim, and the `.cmd` shim resolves the executable
first so `%*` is passed through untouched.

Separately, the bootstrap aborted whenever the GitHub release API was unreachable —
even with the correct runtime already downloaded and extracted — so a rate limit or
an offline machine stopped the MCP server from starting at all. A failed update
check now falls back to the cached runtime and warns on stderr, and an
already-extracted runtime is resolved before any download is considered, so
reclaiming the cached `.zip` no longer requires the network. With nothing usable
cached, the failure stays loud.
