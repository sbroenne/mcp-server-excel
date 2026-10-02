---
"excelmcp": major
---

Replace `copy`, `copy-values`, and `copy-formulas` with one `range` `copy` action
requiring `paste_kind` (CLI `--paste-kind`, batch JSON `pasteKind`).
Select all, values, formulas, formats, or validation, with native transpose
and skip-blank options. Formatting/validation paste preserves cell content;
content overwrite checks cover the actual expanded/repeated destination and
exclude skipped source blanks. All copy kinds require unmerged rectangular
geometry under either overwrite policy. Return resolved paste bounds and
clear the owned application's copy mode after success or failure.
