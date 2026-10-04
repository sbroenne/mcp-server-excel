# ADR-008: Keep guidance at its source and load it for its purpose

**Status:** Current

## Context and decision

Product operation descriptions come from command metadata and native help.
Optional report-formatting skills supply presentation guidance, not another
operation catalog. Shared documentation is prepared into website and package
outputs rather than independently maintained there.

Contributor instructions use shared root/nested `AGENTS.md` files for current
Copilot, Claude Code, and Codex. ADRs supply architectural reasons; instructions
supply actionable rules. Relevant guides are selected by task instead of
loading the entire documentation set.

## Reasons and tradeoffs

Copied catalogs and per-client rule sets can disagree after a change. Keeping
one source limits that drift and avoids filling agent context with instructions
unrelated to the task.

The product skill comparison described in the skill guide supports narrowing
broad automatic guidance for its tested tasks; it does not prove that formatting
skills or this contributor layout improve all agents.

Shared instructions trade client-specific tuning for consistency. Native file
discovery still differs between clients, so a common filename does not guarantee
every nested file is loaded. Generated outputs likewise need preparation before
source edits reach installed packages.

## Implementation and guidance

- [Skill preparation](../scripts/Build-AgentSkills.ps1) and [skill scope](AGENT-SKILLS.md)
- [Website source handling](../gh-pages/sitegen/sources.py)
- [Contributor discovery](agents/development.md)
- [Instruction maintenance](agents/rules/meta.md) and [product guidance sources](agents/rules/mcp-llm-guidance.md)
