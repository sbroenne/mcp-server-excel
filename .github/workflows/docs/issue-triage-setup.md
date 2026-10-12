# Automatic issue triage and small fixes

`issue-triage.md` is the source; `issue-triage.lock.yml` is the generated GitHub
Actions workflow. It runs when an issue is opened or reopened, including reports
from outside contributors. A maintainer can rerun it for one existing issue.
There is no scheduled backlog sweep, edit trigger, or comment trigger.

## What it does

The read-only Copilot triage step checks repository evidence, suggests up to five
existing labels, searches related open and closed issues, and identifies up to
three possible duplicates with reasons. It may request up to three missing
details. It does not apply labels or make a final duplicate decision.

A separate restricted publisher creates one AI-attributed advisory comment.
Reruns update that workflow-owned comment rather than adding more comments.
Unchanged results do not cause a write. Human comments, even those containing
copied automation markers, are never edited.

When automatic fixing is enabled, a clear, localized easy-fix assessment can
assign the existing issue to Copilot. The coding task is instructed to follow
the existing Squad roster, routing, specialist charter, repository rules, and
review requirements. **It does not launch the Squad coordinator itself.**
No Squad profiles, configuration, routing, or state are changed by this workflow.
Existing manually applied `squad`/`squad:*` label workflows remain separate.

The coding agent independently verifies the diagnosis, implements the fix,
and opens or updates one **draft PR** linked to the issue. It must stop if the
investigation no longer supports a small fix. It must not merge, publish, close
the issue, or mark an unchecked fix complete. Human review remains required.

Eligible work can include Excel-related production code, but must have a clear
problem, evidence in current source, at most three suspected existing
implementation files, and a feasible regression check. Directly related tests,
documentation, and changesets may also be needed. The initial scope supports
existing `.cs`, `.ts`, and `.md` files under `src/`, `tests/`, and `docs/`, plus
`README.md` and `FEATURES.md`. Generated output, agent guidance, security-sensitive
paths, scripts, workflows, infrastructure, and dependency changes are excluded.
Possible duplicates, missing details, unclear intended behavior, security or
access-control changes, architecture work, and broad refactoring do not qualify.

Existing assignees or an open PR referencing the issue in its title or body
prevent a new automatic handoff. Security-related wording in the issue, current
labels, non-bot comments, or the assessment summary, reason, or validation
objectives also blocks assignment, including token/key, OAuth, access-control,
and vulnerability reports. This conservative wording check can block harmless
reports too; it is not proof that every security issue can be recognized.
The workflow does not replace
human ownership. Per-issue concurrency
serializes its reruns, and a bot-owned pending receipt is saved before assignment.
A lost API response leaves that receipt pending rather than blindly starting a
second task. Issue edits, human comments, ownership changes, and default-branch
changes during inference invalidate the original assessment.

## Credentials and switches

The hosted triage workflow uses **Actions** configuration:

| Name | Purpose |
| --- | --- |
| `COPILOT_GITHUB_TOKEN` secret | Copilot inference for triage and threat detection |
| `COPILOT_ASSIGN_TOKEN` secret | Existing owner-authenticated coding-agent assignment; available only to the restricted handoff step |
| `ISSUE_AUTO_FIX_ENABLED` variable | Set exactly `true` to permit automatic handoffs; otherwise triage is advisory |
| `EXCEL_RUNNER_ENABLED` variable | Existing runner control must be enabled for automatic handoff |

The built-in workflow token reads issue data and publishes the advisory comment.
It cannot start a coding-agent task. Assignment uses the separately configured
user token, restricted to this repository. GitHub's issues assignment API requires
metadata read and actions, contents, issues, and pull requests read/write for a
fine-grained token. The token must belong to `sbroenne` so issue assignment uses
the existing owner-only runner wake/control path. Authentication failures are
reported, not retried with a different token or weaker assignment method.

Do not put secret values in chat, command arguments, files, logs, or artifacts.
Use GitHub's secret UI or an interactive secret prompt for separately authorized
credential setup. Existing secret names were observed; that alone does not prove
their permissions, expiration, or assignment usability.

To enable new automatic handoffs after delivery and an authorized live check:

```powershell
gh variable set ISSUE_AUTO_FIX_ENABLED --body true --repo sbroenne/mcp-server-excel
```

To stop new automatic handoffs without stopping advisory triage:

```powershell
gh variable set ISSUE_AUTO_FIX_ENABLED --body false --repo sbroenne/mcp-server-excel
```

This does not cancel already assigned coding tasks. To stop new triage runs too:

```powershell
gh workflow disable issue-triage.lock.yml --repo sbroenne/mcp-server-excel
```

Public issue triggers incur inference cost even for reports that do not qualify
for fixing. Automatic coding tasks also incur inference and existing runner
costs. Disabling auto-fix does not disable triage inference.

## Existing Excel runner: two separate variable stores

The existing runner connection is already working. These different values are
intentional, not stale duplicates:

| Store | Variable | Intended value |
| --- | --- | --- |
| Agents | `EXCEL_COPILOT_ENABLED` | `true`: cloud coding sessions use `excel-copilot` |
| Actions | `EXCEL_COPILOT_ENABLED` | `false`: ordinary setup workflow checks use hosted Windows |
| Actions | `EXCEL_RUNNER_ENABLED` | `true`: existing control, maintenance, health, and validation are enabled |

Copilot cloud sessions receive **Agents** variables, not ordinary **Actions**
variables. `gh variable list` inspects Actions and cannot establish cloud-agent
runner selection. Read the applicable store explicitly:

```powershell
gh api repos/sbroenne/mcp-server-excel/agents/variables/EXCEL_COPILOT_ENABLED `
    -H 'X-GitHub-Api-Version: 2026-03-10' --jq '{name,value}'
gh api repos/sbroenne/mcp-server-excel/actions/variables/EXCEL_COPILOT_ENABLED `
    --jq '{name,value}'
```

The authoritative setup and admission explanation is
[Excel runner enablement and acceptance](../../../infrastructure/azure/README.md#enablement-and-acceptance).
Use dedicated Agents settings, not the legacy Actions `copilot` environment.
No runner-selection or firewall setting changes are part of this workflow.
The workflow has no administration permissions to inspect or change Agents
configuration; prerequisite inspection is an operator responsibility.

The actual job for issue #1118 was observed running on `azure-excel-copilot`
and had opened draft PR #1128. That establishes runner routing, not successful
Excel tests or this new workflow's end-to-end assignment.
Keep the existing protected one-job admission, desktop qualification, maintenance,
cleanup, and wake/park controls. Do not execute the triage AI on the desktop.
Runtime fixes require focused regression evidence and sequential
`scripts\Test-E2E.ps1` against final PR source after the last runtime change.
Missing or failed checks must be stated in the draft PR, never reported as passed.

## Manual rerun and troubleshooting

After the workflow is delivered to the default branch:

```powershell
gh workflow run issue-triage.lock.yml --repo sbroenne/mcp-server-excel `
    -f issue_number=123
```

Use an actual open issue number, not a PR number. Closed or locked issues are
skipped with a visible explanation. Invalid inputs and failed required reads
fail visibly. Comments and assignment never use an AI-selected target repository.
The published comment indicates whether handoff is disabled, blocked, pending,
or assigned when the handoff step runs.

If a receipt is pending, inspect the issue assignees, coding sessions, and linked
PRs before any manual retry. Removing an assignee does not reset the receipt.
A maintainer must verify no task was started before repairing an uncertain
receipt; normal reruns intentionally do not restart it. The workflow does not
automatically clean up or cancel partial work.

For stale assessments, rerun on current issue content. For assignment failures,
check the owner token and repository's Copilot availability. For queued desktop
work, use the existing runner control evidence; do not restart a busy runner.
The triage workflow does not create fallback failure issues or extra status
comments. Failure details remain in its run logs.

## Source changes and local verification

Use the already pinned compiler release; compile only this workflow:

```powershell
gh aw compile issue-triage --gh-aw-ref v0.89.21 --validate --no-check-update
node --test tests\ExcelMcp.ScriptSafety.Tests\IssueTriage.test.mjs
dotnet test tests\ExcelMcp.ScriptSafety.Tests\ExcelMcp.ScriptSafety.Tests.csproj `
    -c Release --filter "FullyQualifiedName~IssueTriage|FullyQualifiedName~ValidationAreaTests"
```

Review every action/secret manifest change before using the compiler's
`--approve` flag. It approves the compiler manifest, not credentials, repository
settings, assignment, or merge permission. The new assignment credential
reference is isolated to the trusted handoff step; existing released
checkout/setup/artifact/compiler actions are commit-pinned in the lock file.
Commit source and compiler-generated lock together; never hand-edit the lock.

Local fixtures make no GitHub writes or coding-agent assignments. They cover
target/input validation, existing labels and duplicates, comment ownership and
pagination, repeat-safe publishing, stale inputs, ownership/PR blocking, pending
receipts, restricted assignment, CLI callers, and generated workflow boundaries.
Applicable existing ScriptSafety checks are selected for these workflow/helper
paths without selecting unrelated Excel or packaging checks.

Local tests and compilation cannot prove live inference quality, credential
usability, coding-agent instruction adherence, or complete runtime validation.
An authorized rollout must verify one eligible issue produces one coding session
on the existing runner and one draft PR with actual checks, then verify a rerun
does not start another task. No live assignment or remote setting was changed
while implementing this workflow.

## References

- [Copilot cloud-agent issue assignment API](https://docs.github.com/en/copilot/how-tos/use-copilot-agents/cloud-agent/use-cloud-agent-via-the-api#using-the-issues-api)
- [Dedicated Agents secrets and variables](https://docs.github.com/en/copilot/how-tos/copilot-on-github/customize-copilot/customize-cloud-agent/configure-secrets-and-variables)
- [Restricted gh-aw publishing jobs](https://github.github.com/gh-aw/reference/custom-safe-outputs/)
