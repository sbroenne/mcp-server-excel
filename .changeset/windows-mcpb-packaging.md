---
"excelmcp": patch
---

Fix Windows release packaging by creating a ZIP archive before applying the `.mcpb` extension and shutting down transient MSBuild workers between VSIX clean and publish steps.
