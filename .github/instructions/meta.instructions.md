---
applyTo: ".github/**/*.md,.github/instructions/**,vscode-extension/.github/**/*.md,AGENTS.md,CONTEXT.md,docs/agents/**/*.md"
excludeAgent: "code-review"
---

# Instruction maintenance

- Keep only repo-specific constraints, non-obvious pitfalls, required checks,
  and source pointers. Omit generic coding advice, tutorials, and inventories.
- One authoritative home per rule. Put task-specific rules in scoped
  `*.instructions.md` with quoted `applyTo`; root guidance stays short.
- Scope includes owning generators/templates. Link nested extension guidance
  from the root so it is discoverable.
- Implementation files exclude `code-review`; the review file excludes
  `cloud-agent` and retains essential checks needed independently.
- Consolidation must preserve unique safeguards and update inbound links.
  Keep procedures in developer docs, not automatically loaded instructions.
