# Excel CLI discovery and report-formatting skills

The small CLI discovery skill tells agents how to launch the installed plugin
for ordinary workbook requests. The optional formatting skills supply
presentation conventions for requested reports, not general command catalogs.
Native CLI help and MCP tool schemas remain the source for actions, parameters,
defaults, and safety.

| Skill | Entry point | Distribution |
|-------|-------------|--------------|
| `excel-cli` | Plugin's argument-safe npx launcher | `excel-cli` plugin, skill ZIP (requires the plugin launcher) |
| `excel-cli-report-formatting` | `excelcli` | `excel-cli` plugin, standalone skill ZIP |
| `excel-mcp-report-formatting` | Excel MCP tools | `excel-mcp` plugin, VS Code extension, standalone skill ZIP |

The MCPB configures the server; it does not install agent skills.
General workflows, limitations, and recovery guidance live in
[the documentation reference](reference/README.md), not in every skill package.
All optional [report-formatting conventions](reference/report-formatting.md)
remain available, including financial-model colours and dashboard layout.

## Why the scope changed

The completed real-world comparison used public pytest-skill-engineering 1.0.3
and `gpt-6.1-sol`. All 48 matched cases passed independent workbook checks.
The broad skill was read in all 24 treatment cases, but recorded token usage
was 23.4% higher for MCP and 72.4% higher for CLI. Two interrupted attempts had
unknown usage and are excluded from those percentages, not treated as free.
See [the evidence and limitations](../llm-tests/README.md#measure-whether-skills-help).

This supports removing broad automatic loading for the tested tasks/model.
It does not establish that agents can discover the CLI launcher from a plugin
that contains only a formatting skill. A small `excel-cli` discovery skill
restores that entry point without restoring the broad command/reference corpus.
It does not prove that the new formatting skills improve agent performance.
Their actual selection and value require a separate comparison.

## Installation and migration

Plugin names remain `excel-mcp` and `excel-cli`. They are listed in
[Awesome Copilot](https://github.com/github/awesome-copilot), the default
marketplace in current Copilot clients:

```powershell
copilot plugin install excel-mcp@awesome-copilot
copilot plugin install excel-cli@awesome-copilot
```

Alternatively, install from our direct marketplace:

```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

Choose one marketplace per plugin, not both copies of the same plugin. Existing
direct-marketplace installations can stay in place.

The same plugin repository is also a Claude Code marketplace:

```powershell
claude plugin marketplace add sbroenne/mcp-server-excel-plugins
claude plugin install excel-mcp@mcp-server-excel-plugins
claude plugin install excel-cli@mcp-server-excel-plugins
```

After the release containing this change is published, direct skill installation
uses the new identities:

```powershell
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp-report-formatting
```

The VS Code extension registers only `excel-mcp-report-formatting`.
The source repository contains the three actual skill directories; the
prepared release packages add the formatting skills' entry-point-specific reference.
Use a complete prepared package when installing manually.

Existing standalone skill installations are not automatically updated or
removed. Replace the former broad `excel-cli` skill with the small discovery
skill when updating the CLI plugin. Remove the retired broad `excel-mcp` skill
using your client's skill manager; MCP discovery comes from the server configuration.
Do not remove the plugins or their MCP/CLI launch configuration merely because
the skill identities changed. Local source edits do not update installed copies
or publish new packages.

## Authoring and packaging

| Content | Canonical source |
|---------|------------------|
| Skill selection and short entry instructions | `skills/<skill-name>/SKILL.md` |
| Full optional formatting conventions | `docs/reference/report-formatting.md` |
| General workflows and recovery | `docs/reference/*.md` |
| Installation and authoring instructions | `docs/AGENT-SKILLS.md` |
| Tool descriptions and schema metadata | Core interface XML docs/attributes |
| Minimal MCP server instructions | `src/ExcelMcp.McpServer/Program.cs` |

`skills` contains actual skills only. Do not add a general documentation corpus,
copied command catalog, or another broad entry skill. Keep `excel-cli` limited
to discovery of the plugin's npx wrapper and native help; it carries no references.

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release
.\scripts\Build-AgentSkills.ps1 -GenerateOnly
```

Preparation writes complete skills to `artifacts\generated-skills`.
`Build-AgentSkills.ps1` copies CLI discovery without references. For the formatting
skills it selects only `report-formatting.md`, renders its matching
`cli` or `mcp` fenced examples, and links supporting topics to the website.
The ZIP, plugins, and extension consume the same prepared output. Ordinary
M, DAX, JSON, and other language fences are preserved; no flag translation occurs.

Do not edit generated, packaged, installed, or published copies. Validate both
formatting skill directories and the CLI discovery skill with the public loader
before paid tests.
Availability is not evidence of loading, and loading is not evidence of benefit.
Keep unrelated reads, raw exports, data edits, refreshes, and recovery outside
the formatting trigger. Preserve user/template precedence and all optional
conventions. See [evaluation authoring](../llm-tests/README.md#writing-evaluations).
