---
"excelmcp": patch
---

Creating a named range now rejects invalid names without leaving a temporary name, and accepts English formula syntax with non-US regional settings. Equivalent spellings of existing names and unqualified names that resolve to an active-sheet local name are rejected before changing their reference, visibility, or other metadata. Names in different explicit scopes remain distinct. If assigning the formula fails after a new name is created, the error reports that name and its reference when readable, without automatically deleting it.
