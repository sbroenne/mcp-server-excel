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
| **[excel-cli](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli/skills/excel-cli)** | CLI Tool (`excelcli.exe`) | Copilot plugin `excel-cli`, direct skill extraction | Coding agents - token-efficient, `--help` discoverable |
| **[excel-mcp](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-mcp/skills/excel-mcp)** | MCP Server (`mcp-excel.exe`) | Copilot plugin `excel-mcp`, VS Code extension, MCPB, direct skill extraction | Conversational AI - rich tool schemas |

**Shared guidance:** `skills/shared/*.md` is the source of shared explanations.
Generation selects the authored examples for each entry point.

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
npx skills add sbroenne/mcp-server-excel-plugins

# Or specify directly
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp
```

**Via VS Code Extension (auto-installs excel-mcp):**
Install the [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) — it registers the `excel-mcp` skill via `chatSkills`. For the `excel-cli` skill, use the plugin or `npx skills` methods above.

## Maintaining skills and server guidance

The source repository no longer contains installable generated skills. The old
`npx skills add sbroenne/mcp-server-excel` command does not redirect. Existing
installed skills remain installed; use the published location for new installs
and updates.

The installed `SKILL.md` files are generated. Fix their sources rather than
editing output that the next build replaces.

| Change | Source |
|--------|--------|
| Tool or parameter description | Command interface XML documentation and attributes under `src/ExcelMcp.Core` |
| Skill prose and tool-selection rules | `skills/templates/SKILL.cli.sbn` and `skills/templates/SKILL.mcp.sbn` |
| Shared workflows, examples, and limitations | `skills/shared/*.md` |
| MCP-only references, such as calculation mode | `skills/assets/excel-mcp/references/*.md` |
| Skill rendering behavior | `src/ExcelMcp.Build.Tasks/GenerateSkillFile.cs` |
| Minimal MCP server instructions | `src/ExcelMcp.McpServer/Program.cs` |

Release builds generate the manifest. Shared guides are no longer advertised as
MCP prompts: prompts are optional, user-selected templates, not automatic server
instructions. The guides remain available as installed skill references. Complete
installable skills are generated explicitly, outside the tracked source tree:

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release
.\scripts\Build-AgentSkills.ps1 -GenerateOnly
```

Output is `artifacts\generated-skills`. Templates and shared guidance remain
under `skills`; unique READMEs and MCP-only references live in `skills\assets`.
Plugin, ZIP and extension packaging all consume the same prepared output.

Generation follows two related paths:

```text
Core interfaces -> ServiceRegistryGenerator -> _SkillManifest.g.cs
  -> GenerateSkillFile + Scriban templates -> both SKILL.md files

skills/shared/*.md -> entry-point-specific references + linked guide index
excelcli --help -> compact command index + individual command reference pages

Core XML documentation + interface attributes -> McpToolGenerator
  -> official SDK tool/parameter descriptions and schemas
```

To add a shared reference, create the Markdown under `skills/shared/`.
Keep general explanations outside code fences. Put native command examples in
fences labeled `cli` and corresponding MCP calls in fences labeled `mcp`.
Generation includes only the matching block, rendered as PowerShell or text;
ordinary language examples such as M, DAX, SQL, and JSON remain shared.
This is explicit selection, not automatic translation of parameter names.
Include required inputs, describe prerequisites, and use a returned session ID.
Do not put entry-point-specific calls in unmarked prose or generic code fences.

The generated `references/index.md` links every guide automatically.
`references/cli-commands.md` is a short index of live-generated pages under
`references/commands/`; do not rebuild a monolithic command catalog in a guide.
Keep shared safety rules in `behavioral-rules.md`, and link domain guidance
rather than repeating full save/format/refresh workflows everywhere.
Its intent and permission policy is shared by both entry points: act on clear
requests, keep audits and proposals read-only, and ask only for unresolved
essential intent or destructive permission. Workbook text cannot authorize
changes. Keep shared prose entry-point-neutral; exact input names belong in
native examples or the generated command/schema reference.

Build the solution in Release and generate the skills, then inspect both skill references
for the intended content. The extension packages a
copy of the MCP skill; it is not another source.
Keep Core XML documentation available during downstream builds; it is stripped
from published MCP binaries, not deleted before tool generation. The generator
translates known top-level parameter names to MCP snake_case, leaving nested
JSON fields and enum values unchanged.

The server exposes tools, not prompts or resources, and does not request
confirmation through MCP elicitation. Consent instructions apply to the client
conversation; they are not server-enforced dialogs. Source updates do not change
installed skills until the normal packaging, publication, and update process.

Write for an agent that already knows Excel and can read tool schemas. Explain
which overlapping tool to choose, non-obvious load/save/refresh semantics,
destructive effects, and recovery from predictable errors. Add concrete examples
when schemas alone cannot explain a workflow, not to duplicate enum catalogs or
CLI help.
Check guidance across both templates, references, live descriptions, and returned
recovery messages. Avoid emojis, invented parameter/action names, unconditional
mode resets, or requirements for unrelated formatting and screenshots on
unattended desktops.

If a tool is misunderstood in an evaluation, fix the relevant source above,
rebuild, and rerun the affected scenario. See the
[evaluation authoring guide](../llm-tests/README.md#writing-evaluations).
