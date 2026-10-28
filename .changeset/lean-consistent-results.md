---
"excelmcp": minor
---

CLI commands and batch items now share one JSON layout, with returned values under `data`, failure details under `error`, and an operation time in `durationMs`. Scripts must update their field lookups; batch items keep their `index` but no longer repeat a result wrapper or command summary.

Important warnings, failures, partial state, and inspection limits remain visible. Power Query lists omit formula previews; full query reads remain unchanged.
