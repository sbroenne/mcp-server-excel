---
"excelmcp": patch
---

Correct MCP tool schemas for Gemini clients: preserve nullable array annotations
and describe mixed cell values as strings, numbers, booleans, or null instead of
claiming all cells are converted to strings. Keep nested arrays compatible with
legacy adapters by avoiding array-valued `type` keywords.
