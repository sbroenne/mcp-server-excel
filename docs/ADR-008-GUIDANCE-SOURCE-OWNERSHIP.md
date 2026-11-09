# ADR-008: Keep guidance at its source and load it for its purpose

**Status:** Current

## Context and decision

Product operation descriptions come from command metadata and native help.
Only the CLI plugin bundles a short `excel-cli` quick-start skill for the session,
batch and verification routine. Detailed CLI discovery uses native help and
`list-actions`; plugin documentation identifies the npx launcher. Optional report
formatting advice stays in the reference documentation, not automatically loaded
agent instructions. Shared documentation is prepared into website outputs rather
than independently maintained there.

Contributor instructions use shared root/nested `AGENTS.md` files for current
Copilot, Claude Code, and Codex. ADRs supply architectural reasons; instructions
supply actionable rules. Relevant guides are selected by task instead of
loading the entire documentation set.

## Reasons and tradeoffs

Copied catalogs and per-client rule sets can disagree after a change. Keeping
one source limits that drift and avoids filling agent context with instructions
unrelated to the task.

The earlier, longer skills showed no measurable benefit and were removed.
A later small trial of a shorter CLI quick start reduced help lookups, tokens
and elapsed time. This supports a narrow CLI-only exception, not restoring
duplicate formatting guidance or claiming a general improvement across models
and tasks. The quick start does not change supported workbook operations.

Shared instructions trade client-specific tuning for consistency. Native file
discovery still differs between clients, so a common filename does not guarantee
every nested file is loaded. Generated outputs likewise need preparation before
source edits reach installed packages.

## Implementation and guidance

- [Native agent guidance](AGENT-SKILLS.md)
- [Website source handling](../gh-pages/sitegen/sources.py)
- [Contributor discovery](agents/development.md)
- [Instruction maintenance](AGENTS.md#instruction-maintenance) and [product guidance sources](agents/product-guidance.md)
