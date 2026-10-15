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
  checks near the beginning of root `AGENTS.md` under `Code Review Rules`, so
  loading the shared instructions includes the critical checks early. Do not
  maintain a second checklist for an unused review client.
- Audit rules against source, executable checks, and vendor documentation.
  Correct or remove stale rules; code violating a rule is not by itself proof
  that the rule is obsolete. Preserve valid safeguards and update inbound links.
  Remove repeated pattern instructions only when an implemented compiler check
  covers them; keep ownership, runtime invariants, and validation procedures.
  A suppression justification is evidence to review, not proof of safety.
- ADRs explain current choices and tradeoffs; instructions tell agents what
  to do. Link between them rather than copying rules; ADR authoring rules live
  in [maintaining decisions](../../DECISIONS.md#maintaining-decisions).
  Keep procedures in developer docs, not automatically loaded instructions.
