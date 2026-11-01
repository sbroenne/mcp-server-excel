---
"excelmcp": patch
---

Keep numbers that happen to match Excel error codes as numbers in range reads and formula traces. Formulas returning these numbers are also no longer reported as blocked spills; actual cell errors still include their error names and diagnostics.
