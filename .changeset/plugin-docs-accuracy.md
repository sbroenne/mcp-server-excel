---
"excelmcp": patch
---

**Accurate plugin documentation and a VERSION file for the CLI plugin.** The MCP
server's `--help` banner claimed "22 tools with 195+ operations" while the server
actually registers 31 tools with 326 operations. The repo already derives those
numbers from code and enforces them across 16 documents on every commit; the banner
was simply not one of them, so it drifted unnoticed.

Rather than correcting the literal, the banner now *derives* both numbers by
reflecting over the live `[McpServerTool]` registration, so it can no longer disagree
with the server's own `tools/list` response. The doc-count guard was extended to fail
if anyone reintroduces a hard-coded count there — including a count that happens to be
correct on the day it is written.

The `excel-cli` plugin shipped without the `VERSION` file its `excel-mcp` counterpart
carries, because the build never passed a version through for the CLI skill and only
rewrote a `VERSION` that already existed instead of creating one. Both plugins now
get a stamped `VERSION`, and the build fails if any packaged skill is missing one or
carries the wrong version.

Two skill instructions were misleading in ways that produce visibly wrong output. The
number-format table showed rendered results as though separators were fixed, but
Excel renders them per the user's Windows regional settings — `$#,##0.00` shows
`$1.234,56` on a German machine — so the skill now explains that format codes are
written in US notation while the rendering is locale-dependent, and warns against
"fixing" the code. The formatting workflow also stopped at applying a number format,
which leaves date and currency columns showing `#####` because formatted values are
wider than the raw ones; auto-fitting columns is now a required step.

Finally, the CLI skill assumed `excelcli` was on PATH while the plugin's global shim
is explicitly opt-in, so an agent following the skill hit command-not-found. The
preconditions now state the requirement plainly and give the ways to satisfy it.
