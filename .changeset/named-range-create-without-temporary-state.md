---
"excelmcp": patch
---

Creating a named range now rejects invalid names without leaving a temporary name, and accepts English formula syntax with non-US regional settings. If assigning the formula fails after the requested name is created, the error reports the remaining name and its reference when readable, without automatically deleting it.
