---
"excelmcp": minor
---

Add selected-series reads and primary/secondary axis assignment, native error
bars, per-point material and marker formatting, and actual chart image export
through MCP and CLI. Fix secondary axis titles and number formats targeting a
primary axis. Report unavailable native getters explicitly and reject unsupported
marker transparency/outline-weight edits and PivotChart per-series mutations.
Failed image exports remove new empty or partial output while preserving an
existing destination image.
