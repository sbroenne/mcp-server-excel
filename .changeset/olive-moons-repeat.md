---
"excelmcp": patch
---

**Localized Excel Table formats survive save**: Opening and saving a workbook no
longer rewrites locale-specific number formats on Excel Table columns. Japanese
(ja-JP) built-in date and negative-number formats now stay exactly as they were.
