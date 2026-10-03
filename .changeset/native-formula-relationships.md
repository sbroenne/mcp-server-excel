---
"excelmcp": minor
---

Add native precedent/dependent graph inspection to MCP and CLI. Return all
reachable native worksheet relationships, current formulas/values, cycles,
and unresolved lookups without changing selection or opening external files.
Explicit coverage explains missing cross-sheet, external, and dynamic references;
an ambiguous native lookup is never reported as proven empty coverage.
