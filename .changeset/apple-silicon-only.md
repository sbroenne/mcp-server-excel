---
"excelmcp": patch
---

Limit native macOS distribution to Apple Silicon across npm, standalone ZIP,
VSIX, Copilot plugin, and MCPB packages. Intel Macs now fail as unsupported
instead of selecting or downloading an x64 runtime.
