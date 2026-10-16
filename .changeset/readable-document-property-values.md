---
"excelmcp": patch
---

Document-property reads now report unreadable built-in values through `valueAvailable` and `valueError` instead of failing the entire listing. Property names, types, and scopes remain available, and a failed value read is not mistaken for proof that the value is undefined. Other unexpected native failures still fail the command.
