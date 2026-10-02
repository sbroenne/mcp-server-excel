# Instruction maintenance

- Keep only repo-specific constraints, non-obvious pitfalls, required checks,
  and source pointers. Omit generic coding advice, tutorials, and inventories.
- One authoritative home per rule. Common coding rules belong in root
  `AGENTS.md`; use nested `AGENTS.md` files and neutral shared guides for
  task-specific rules. Maintain the root task/path map; root guidance stays short.
- Scope includes owning generators/templates. Link nested extension guidance
  from the root so it is discoverable.
- Distinguish implementation tasks from review tasks explicitly. Keep the
  standalone review checklist in `.github/copilot-instructions.md` for VS Code
  built-in review, and link it from `AGENTS.md` for other agents. Do not create
  another maintained copy or rely on a pointer alone for built-in review.
- Consolidation must preserve unique safeguards and update inbound links.
  Keep procedures in developer docs, not automatically loaded instructions.
