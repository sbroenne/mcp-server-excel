---
"excelmcp": patch
---

Creating a named range no longer reports success while leaving a temporary name when Excel rejects the requested name. Creation now uses the requested name and formula in one native call, without automatic deletion after a failure.
