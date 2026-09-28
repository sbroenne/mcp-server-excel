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
helper add-in. Helper version 1.1.0 includes fixed, transactional candidates for
Power Query authoring, synchronous refresh, worksheet load transitions, and
temporary-query evaluation. Installing the helper does not enable unproven
actions; every new method remains individually gated by prompt-free CLI and MCP
evidence.
