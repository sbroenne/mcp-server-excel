---
"excelmcp": minor
---

Replace `range_format` / `rangeformat` actions `format-range` and `format-ranges`
with one `format` action taking `rangeAddresses` and typed `formatOptions`.
The old actions and scalar formatting inputs are removed.

Add independent edge, inside, and diagonal borders; theme colors and tints;
native underline kinds, font effects and theme fonts; indentation, shrink-to-fit,
reading order, fill alignment, and center-across-selection. Validate all targets,
protection, and known invalid options before writing. Omitted settings preserve
native state; native failures do not promise rollback.
