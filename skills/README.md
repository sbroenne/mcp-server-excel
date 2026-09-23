# Excel MCP Server - Agent Skills

**Skills teach your AI assistant how to use Excel MCP Server well.** A skill is a
small package of guidance and examples that your coding agent (GitHub Copilot,
Cursor, Windsurf, Claude Code, and others) loads automatically — so it knows the
right workflow, the correct parameters, and the common gotchas without you
having to spell them out each time. Installing a skill makes the assistant
noticeably more reliable at driving Excel.

There are two packages — pick the one that matches how you connect to Excel (or
install both):

| Skill | Component | Distribution | Best For |
|-------|-----------|--------------|----------|
| **[excel-cli](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-cli/SKILL.md)** | CLI Tool (`excelcli.exe`) | Copilot plugin `excel-cli`, direct skill extraction | Coding agents - token-efficient, `--help` discoverable |
| **[excel-mcp](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-mcp/SKILL.md)** | MCP Server (`mcp-excel.exe`) | Copilot plugin `excel-mcp`, VS Code extension, MCPB, direct skill extraction | Conversational AI - rich tool schemas |

**Shared guidance:** `skills/shared/*.md` — source of truth for both skills (auto-copied to each skill's `references/` folder)

> **Note:** Legacy npm packages (`excel-cli-skill`, `excel-mcp-skill`) are no longer published. Use the methods below instead.

## Installation

**GitHub Copilot Plugins (Recommended):**
```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

**Direct skill extraction (for agents without plugin support):**
```powershell
# Via npx (interactive — select excel-cli, excel-mcp, or both)
npx skills add sbroenne/mcp-server-excel

# Or specify directly
npx skills add sbroenne/mcp-server-excel --skill excel-cli
npx skills add sbroenne/mcp-server-excel --skill excel-mcp
```

**Via VS Code Extension (auto-installs excel-mcp):**
Install the [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) — it registers the `excel-mcp` skill via `chatSkills`. For the `excel-cli` skill, use the plugin or `npx skills` methods above.

## Maintaining skills and MCP prompts

The installed `SKILL.md` files are generated. Fix their sources rather than
editing output that the next build replaces.

| Change | Source |
|--------|--------|
| Tool or parameter description | Core interface XML documentation and attributes |
| Skill prose and tool-selection rules | `templates/SKILL.cli.sbn` and `templates/SKILL.mcp.sbn` |
| Shared workflows, examples, and limitations | `shared/*.md` |
| Skill rendering behavior | `src/ExcelMcp.Build.Tasks/GenerateSkillFile.cs` |
| MCP prompt description overrides | `GenerateSkillPromptsClass` in the MCP Server project file |

Release builds follow two related paths:

```text
Core interfaces -> ServiceRegistryGenerator -> _SkillManifest.g.cs
  -> GenerateSkillFile + Scriban templates -> both SKILL.md files

skills/shared/*.md -> copied skill references
  -> embedded MCP prompt content + generated ExcelSkillPrompts.g.cs
```

To add a shared reference, create the Markdown under `shared/`. Review the MCP
prompt description overrides if its automatic description is insufficient.
Build the solution in Release, then inspect both skill references and the
generated prompt surface for the intended content. The extension packages a
copy of the MCP skill; it is not another source.

Write for an agent that already knows Excel and can read tool schemas. Explain
which overlapping tool to choose, non-obvious load/save/refresh semantics,
destructive effects, and recovery from predictable errors. Add concrete examples
when schemas alone cannot explain a workflow, not to duplicate enum catalogs or
CLI help.

If a tool is misunderstood in an evaluation, fix the relevant source above,
rebuild, and rerun the affected scenario. See the
[evaluation authoring guide](../llm-tests/README.md#writing-evaluations).
