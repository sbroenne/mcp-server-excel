---
applyTo: ".github/**/*.md,.github/instructions/**"
excludeAgent: "code-review"
---

# Maintaining Copilot instructions

- Repository-wide guidance belongs in `.github/copilot-instructions.md` and
  should stay short, durable, and task-independent.
- Path guidance uses `NAME.instructions.md` under `.github/instructions` with a
  quoted `applyTo` glob in YAML frontmatter.
- Add `excludeAgent: "code-review"` to implementation guidance that should not
  influence review; the dedicated review file excludes `cloud-agent`.
- Put a rule at the narrowest useful scope. Do not repeat it in repository-wide,
  critical, and feature files.
- Describe the current architecture and commands. Remove incident history,
  obsolete tool names, dated counts, and one-off examples.
- Use natural language and real repository paths. Do not prescribe a specific
  assistant's private tool names.
- When implementation and instructions diverge, verify the code and update the
  stale instruction in the same change.
