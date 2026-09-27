---
"excelmcp": minor
---

Add a bounded, non-prompting macOS VBA capability preflight. VBA requests now
report whether effective Office preferences disable macros, require workbook
approval, permit unattended execution, or cannot be determined, and separately
report user-managed VBA project-model trust. ExcelMcp never changes those
preferences. Macro execution and source CRUD remain gated until repository-owned
fixtures and a safe project-model route satisfy their independent evidence
requirements. The distribution now also includes reviewable source and a
strict, versioned transport foundation for an optional user-installed Mac Excel
helper add-in. Installing the helper does not enable unproven actions.
