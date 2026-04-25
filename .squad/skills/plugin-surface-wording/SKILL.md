---
name: "plugin-surface-wording"
description: "Describe plugin release mechanics without overclaiming client-specific installation behavior."
domain: "release-management"
confidence: "high"
source: "earned"
---

## Context
Use this when documenting a repo that publishes plugin bundles consumed by one or more agent/plugin clients.

## Patterns
- Describe the release output as plugin artifacts, plugin bundles, or agent plugins.
- Separate artifact publication from client-specific installation UX and marketplace registration.
- Keep install commands only for the clients you have actually verified and documented.
- If one client is the only verified install path today, label those commands as examples for that client instead of implying exclusivity or universal compatibility.

## Examples
- "Publishes plugin artifacts to the published repo" is safer than "publishes to every plugin marketplace."
- "Copilot CLI installation examples" accurately scopes commands that are only verified for Copilot.

## Anti-Patterns
- Calling a plugin package CLI-only when the bundle format can also matter to other plugin-capable clients.
- Claiming VS Code or Claude marketplace publication when the workflow only pushes files to a GitHub repo.
