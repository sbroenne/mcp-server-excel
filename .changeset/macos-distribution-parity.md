---
"excelmcp": minor
---

Harden macOS distribution with architecture-isolated ARM64 and x64 npm, ZIP,
VSIX, and MCPB packages; clean-install archive checks; executable-mode
validation; Developer ID signing; and Apple notarization hooks. Launchers fail
closed instead of selecting a runtime for the wrong architecture. Intel
packages are cross-built and structurally validated; physical Intel Mac Excel
execution remains unverified.
