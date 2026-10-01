---
"excelmcp": patch
---

MCP session identifiers now consistently use `session_id` in list entries and errors as well as open/create results and inputs. Calls using the old `sessionId` input are rejected; CLI output naming is unchanged. Tool descriptions and both Excel skills now explain safe saving, calculation modes, dependent-call ordering, and destructive actions with no tool-level undo.

Destructive-action guidance no longer prescribes unsolicited workbook copies. Calculation guidance now states that restoring the prior mode is best-effort and can fail without failing the write.
