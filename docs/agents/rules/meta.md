# Instruction maintenance

- Keep only repo-specific constraints, non-obvious pitfalls, required checks,
  and source pointers. Omit generic coding advice, tutorials, and inventories.
- One authoritative home per rule. Common coding rules belong in root
  `AGENTS.md`; use nested `AGENTS.md` files and neutral shared guides for
  task-specific rules. Maintain the root task/path map; root guidance stays short.
- Target native `AGENTS.md` support in current Copilot, Claude Code, and Codex.
  Do not copy rules into per-client files or add older-version import wrappers.
  Document client discovery conditions in [agent development](../development.md),
  not as assumptions that every nested file or Markdown link is auto-loaded.
- Scope includes owning generators/templates. Link nested extension guidance
  from the root so it is discoverable.
- Distinguish implementation tasks from review tasks explicitly. Keep review
  checks directly in root `AGENTS.md` under `Code Review Rules`, so loading the
  shared instructions includes the checks without following a link. Do not
  maintain a second checklist for an unused review client.
- Audit rules against source, executable checks, and vendor documentation.
  Correct or remove stale rules; code violating a rule is not by itself proof
  that the rule is obsolete. Preserve valid safeguards and update inbound links.
- ADRs explain current choices and tradeoffs; instructions tell agents what
  to do. Link between them rather than copying rules, checklists, or procedures.
  Maintain current decisions only; Git retains earlier versions.
  Keep procedures in developer docs, not automatically loaded instructions.
