---
"excelmcp": patch
---

**Clear Python in Excel availability errors** (#753): `pythoninexcel set-formula` and `get-result` now explain when the current Excel session cannot use Python in Excel instead of reporting success or exposing a raw `#NAME?` worksheet error.
