# Development with coding agents

[AGENTS.md](../../AGENTS.md) owns the common repository rules and links the
task-specific guides. Read [CONTEXT.md](../../CONTEXT.md) and the current ADRs
before working on an unfamiliar area.

## Instruction discovery

| Client | Repository instructions |
| --- | --- |
| Copilot CLI (`copilot`) and cloud agent | Root AGENTS.md and relevant nested guidance |
| Copilot app | AGENTS.md; accepted `.github/github-app.yml` adds pointers and manual commands |
| VS Code Copilot chat | Root AGENTS.md with `chat.useAgentsMdFile` enabled |
| VS Code Local agent, nested guidance | `chat.useNestedAgentsMdFiles` enables discovery; it is disabled by default |
| VS Code Agent Host sessions | Follow the selected agent's discovery rules |
| VS Code built-in Copilot review | Standalone `.github/copilot-instructions.md` review checklist |
| GitHub.com Copilot PR reviews | AGENTS.md and the same standalone review checklist |
| Claude Code | AGENTS.md natively in supported versions, subject to the conditions below |
| GitHub CLI (`gh`) | Repository administration commands; not an instruction-reading coding agent |

The root task map explicitly tells agents which guides to read, including when
nested discovery is disabled. Instructions are guidance, not technical
enforcement or permission to publish.

Claude Code's native AGENTS.md support requires v2.1.277 or later and the
built-in `agents-md` plugin. By default, a project or ancestor `CLAUDE.md`,
`.claude/CLAUDE.md`, or `CLAUDE.local.md` takes priority instead. User-wide
`~/.claude/CLAUDE.md` does not prevent the project fallback. Check the
**Project instructions** setting and loaded files in your session.
For a session without native support, a `CLAUDE.md` containing `@AGENTS.md`
imports the shared file without maintaining another rule book. Do not add
a plain-text pointer and assume it is automatically imported.
The video subproject retains its explicit import of its own AGENTS.md.

Other agents can use the same ordinary Markdown guides, but not every client
automatically reads AGENTS.md. Do not assume that a file being present proves
the client loaded or followed it.

## Windows setup

Install PowerShell 7, the .NET SDK selected by `global.json`, Node.js 22,
Python 3.13, and uv. Use the checked-in lockfiles:

```powershell
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
dotnet restore Sbroenne.ExcelMcp.sln
npm ci
npm --prefix vscode-extension ci
npm --prefix npm-packages/shared ci
python -m pip install -r gh-pages\requirements.txt
uv sync --project llm-tests --locked
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -Contracts
```

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

## Sources

- [GitHub instruction support](https://docs.github.com/en/copilot/reference/custom-instructions-support)
- [VS Code instruction discovery](https://code.visualstudio.com/docs/agent-customization/custom-instructions)
- [Claude Code instructions](https://code.claude.com/docs/en/memory)
- [App configuration and trust](https://docs.github.com/en/copilot/reference/github-copilot-app-reference/repository-configuration)
- [Cloud setup](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/customize-the-agent-environment)
- [Agent secrets and variables](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/configure-secrets-and-variables)
