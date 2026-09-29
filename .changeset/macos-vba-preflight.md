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
helper add-in. Helper version 1.4.0 includes fixed, transactional candidates for
Power Query authoring, synchronous refresh, worksheet load transitions, and
temporary-query evaluation, plus a bounded read-only observation of late-bound
XML Maps and workbook-model APIs. It also supplies strict public-contract
adapters for VBA source lifecycle and exact workbook-qualified
`Module.Procedure` execution with bounded string parameters. Installing the
helper does not enable unproven actions; every new method remains individually
gated by prompt-free CLI and MCP evidence.

Add a guarded public Power Query lifecycle acceptance runner for helper 1.4.0.
It requires explicit user confirmations, a dedicated Excel-authored workbook,
and an exclusive desktop slot; scopes the exact candidate actions to child CLI
and MCP processes; uses only literal credential-free M; verifies supported
worksheet/connection-only behavior with exact loaded-cell checks across saved
reopens, keeps unconfirmed opens/closes for recovery, verifies gated Data Model
variants, and emits validation-only receipts that cannot be mistaken for
runtime proof.

Add an opaque signed-helper packaging workflow for the planned prebuilt macOS
VBA add-in. It requires an explicit Windows Excel signature verification,
rejects private-key material and invalid SelfCert profiles, and records
whole-file hashes, reviewed source versions, source commit, certificate
identity, fingerprint, and validity. ExcelMcp never signs, opens, trusts, or
modifies the helper during packaging.
