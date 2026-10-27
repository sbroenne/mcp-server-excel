---
"excelmcp": patch
---

ExcelMcp now includes readable text from open Excel prompts in some data connection and PivotTable errors, and when a prompt blocks saving, closing, refresh checks, or account settings; if the text cannot be read, the error says so. It suggests finishing any needed sign-in and retrying before refreshing data, which still needs separate permission because it can change workbook data.

Save As also checks for prompts when working with a workbook that was already open, without answering prompts, reading passwords, or saving or discarding edits.
