---
"excelmcp": minor
---

CLI commands and batch items now share one JSON layout, with returned values under `data`, failure details under `error`, and an operation time in `durationMs`. Scripts must update their field lookups; batch items keep their `index` but no longer repeat a result wrapper or command summary.

Choose returned fields on list/read operations through CLI `--fields`, MCP `fields`, or batch `args.fields`. Important warnings, failures, partial state, and inspection limits remain visible. Power Query lists omit formula previews unless explicitly selected; full query reads remain unchanged.
