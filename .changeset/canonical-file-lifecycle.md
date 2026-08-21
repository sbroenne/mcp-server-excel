---
"excelmcp": major
---

**Canonical file lifecycle** (#798, #799): CLI and MCP now expose the same
list/open/create/close/test workflow. Standalone CLI save and the MCP
`close-workbook` no-op are removed; file testing shares one result model with
openability and deterministic IRM/AIP read-only requirements. IRM detection now
requires the rights-management data-space marker, so ordinary password-encrypted
OOXML files are not incorrectly forced into read-only mode.
