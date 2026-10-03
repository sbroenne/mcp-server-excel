# Development with coding agents

[AGENTS.md](../../AGENTS.md) owns the common repository rules and links the
task-specific guides. [CONTEXT.md](../../CONTEXT.md) provides the system map;
the [decision index](../DECISIONS.md) links architectural reasons and tradeoffs.
Read only the decisions relevant to the work. ADRs are ordinary documentation,
not a special instruction format that agents are guaranteed to load.

## Instruction discovery

Use current Copilot, Claude Code, and Codex versions with native `AGENTS.md`
support. Keep rules in the shared root/nested files, not copied into separate
Copilot or Claude instruction sets. This repository does not provide
older-version compatibility wrappers or a VS Code built-in review configuration.

| Client | Discovery and verification |
| --- | --- |
| Copilot CLI (`copilot`) | Discovers root and applicable nested AGENTS.md files. Use `/instructions` to inspect discovered/enabled files; restart or start a new session after instruction changes. |
| Copilot cloud agent | Supports root/nested AGENTS.md. Check its session evidence rather than assuming every linked guide was read. |
| Copilot app | Uses its selected agent's discovery; accepted `.github/github-app.yml` adds pointers and manual commands. |
| Claude Code | Reads AGENTS.md natively under the conditions below. Use `/context` to inspect memory files and `/config` to check Project instructions. |
| Codex | Builds a root-to-working-directory instruction chain at session startup. Start a fresh session in the intended directory and ask which instruction sources were loaded. |

The root task map explicitly tells agents to read relevant nested/shared guides.
This matters for Codex started at the repository root: its startup walk does not
automatically include instructions in every descendant directory. Codex prefers
`AGENTS.override.md` over `AGENTS.md` in the same directory and limits combined
project instructions to 32 KiB by default. Do not change user configuration to
compensate for unnecessarily large repository instructions.

Claude Code native support requires v2.1.277 or later and the built-in
`agents-md` plugin; use a current version, since early versions had additional
loading exceptions. By default, a project or ancestor `CLAUDE.md`,
`.claude/CLAUDE.md`, or `CLAUDE.local.md` takes priority over AGENTS.md.
User-wide `~/.claude/CLAUDE.md` does not suppress that fallback. If personal
project files affect discovery, the user can select **claude-md-and-agents-md**
in **Project instructions**. Do not edit their settings or add a repository
wrapper automatically.

GitHub also supports `.github/copilot-instructions.md` and path-specific
`.github/instructions/*.instructions.md`. They are not required to duplicate
AGENTS.md. Copilot CLI combines applicable instructions without defining a
general precedence order between all files; do not apply another client's
precedence rules to it. A Markdown link is not an automatic file import.

Instructions guide behavior; they do not enforce permissions or prove adherence.
The checks above inspect loading, not task quality. Shared
[review checks](review.md) apply when requesting a coding-agent review.

## Windows setup

Install PowerShell 7, the .NET SDK selected by `global.json`, and Node.js 22.
Website work also needs Python 3.13. Use the checked-in lockfiles:

```powershell
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
dotnet restore Sbroenne.ExcelMcp.sln
npm ci
npm --prefix vscode-extension ci
npm --prefix npm-packages/shared ci
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -Contracts
```

For website setup, use its [virtual-environment instructions](../../gh-pages/README.md#setup-one-time).
Install uv and the evaluation dependencies only for explicitly requested
[evaluation work](../../llm-tests/README.md); they are not ordinary setup steps.

Use the smallest applicable checks in [AGENTS.md](../../AGENTS.md#build-and-validation).
Builds and Excel-free checks do not prove Excel COM behavior.
Desktop Excel is required for the local tests described in
[tests/README.md](../../tests/README.md). Run Excel-dependent test commands
sequentially. Runtime changes require local Excel E2E; report unavailable
validation rather than claiming the hosted runner covered it.
LLM evaluations also require external model access and may incur costs.
Do not run them as an incidental setup check.

## Cloud and code-review setup

`.github/workflows/copilot-setup-steps.yml` prepares a Windows environment from
the SDK and dependency lockfiles. GitHub-hosted runners do not have Excel.
Only setup on the default branch is adopted by the cloud agent.
If a setup step fails, GitHub may still start the agent after skipping remaining
steps. Diagnose and report missing prerequisites before claiming validation.

The existing review environment shares that setup. Do not add another setup
workflow unless it provides a demonstrated benefit.
GitHub's integrated agent firewall does not protect Windows runners;
this configuration does not change firewall settings or provision runners.

Use GitHub's dedicated **Agents** secrets and variables for current agent
configuration, not the legacy Actions `copilot` environment.
Legacy values were migrated by GitHub; do not copy or inspect their contents.
`COPILOT_MCP_` values are reserved for MCP servers. Publishing secrets for
Actions are separate.

The repository's `.vscode/mcp.json` configures VS Code, not the CLI/app/cloud.
Do not enable ExcelMcp on a hosted runner without Excel or automatically
install it into its own development sessions.
Product formatting skills are prepared from their canonical sources; see
[agent skills](../AGENT-SKILLS.md). Installed and packaged copies
are outputs, not contributor instruction sources.

## App configuration and trust

`.github/github-app.yml` supplies only instruction pointers and manual
**Setup**, **Release build**, and **Excel-free checks** commands.
Release build compiles locally; it does not publish a release.
There are no automatic setup/cleanup triggers, extensions, remote-control
changes, or browser-launch settings.

When opening or reopening this repository, review the configuration's commands
and dependencies before accepting it in the app. Changes made outside the app,
even comments or whitespace, require renewed acceptance. Until acceptance,
the app keeps the previous project settings; UI-authored changes are trusted
automatically. Plain AGENTS.md instructions do not justify trusting unrelated
executable content.

Scripts and their child processes receive GitHub account credentials from the
app. Never log or persist credential environment variables. Setup runs
dependencies, so inspect their provenance before accepting or running it.
Existing UI-only project instructions and the actual trust prompt were not
available through the inspected session API; this file does not establish
that the app's trust flow has been exercised or that earlier UI instruction
text has been preserved. Automation/browser fields are deliberately omitted.

The app's existing **Setup** command includes website and evaluation dependency
installation. It is broader than the ordinary setup above; use the narrower
commands for non-evaluation work. This documentation change does not alter the
app's executable commands or its acceptance state.

## Sources

- [GitHub instruction support](https://docs.github.com/en/copilot/reference/custom-instructions-support)
- [GitHub instruction writing guidance](https://docs.github.com/en/copilot/concepts/prompting/response-customization)
- [Copilot CLI discovery](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/add-custom-instructions)
- [Claude Code instructions](https://code.claude.com/docs/en/memory)
- [Codex instruction discovery](https://developers.openai.com/codex/guides/agents-md)
- [App configuration and trust](https://docs.github.com/en/copilot/reference/github-copilot-app-reference/repository-configuration)
- [Cloud setup](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/customize-the-agent-environment)
- [Agent secrets and variables](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/configure-secrets-and-variables)
