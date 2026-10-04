# Azure Infrastructure

This directory contains the Azure Application Insights infrastructure used for
telemetry development and validation, and the dedicated Excel development VM.

## Files

| File | Purpose |
|---|---|
| `appinsights.bicep` | Application Insights deployment entry point |
| `appinsights.parameters.json` | Production telemetry deployment defaults |
| `appinsights-test.bicep` | Test telemetry deployment entry point |
| `appinsights-test.parameters.json` | Test telemetry deployment defaults |
| `appinsights-resources.bicep` | Shared Application Insights resources and ingestion-time privacy transforms |
| `deploy-appinsights.ps1` | Application Insights deployment helper |
| `configure-analytics-oidc.ps1` | Read-only GitHub Actions workload identity setup |
| `excel-runner.bicep` | Dedicated Windows Excel development VM, network, password vault and shutdown protection |
| `install-excel-office.ps1` | SYSTEM worker for current 64-bit Excel 2024 Retail installation |
| `install-excel-toolchain.ps1` | SYSTEM worker for the repository SDK, Git, PowerShell 7, Node.js 22 and Python 3.13/pip |
| `setup-excel-desktop.ps1` | Dedicated non-admin desktop account, secure automatic logon and profile initialization |
| `excel-desktop-access.bicep` | Free Developer Bastion for temporary private activation access |
| `update-excel-runner.ps1` | Bounded Windows Update and separate Click-to-Run Excel servicing |
| `test-excel-desktop.ps1` | Non-admin account privacy, retail device licence, calculation and saved-workbook checks |
| `configure-excel-runner.ps1` | Transient encrypted registration without starting a listener |
| `start-excel-runner.ps1` | Protected current-boot admission and one complete job per listener |
| `excel-job-started.ps1`, `excel-job-completed.ps1`, `recover-excel-jobs.ps1` | Admission, owned process/workspace cleanup and cancellation recovery |
| `configure-excel-control-oidc.ps1` | Password-free GitHub control identity and default-branch environment restrictions |
| `excel-runner-control.bicep`, `excel-runner-control-assignment.bicep` | Dedicated-group VM control permissions and read-only identity-role auditing |

The workspace transform drops noisy runtime metrics and rejects every
`AppExceptions` row that was not explicitly sanitized by ExcelMcp. Exception
messages and stack traces are never retained; only approved type, source, and
project-owned failure-site classifications are ingested.

## Public usage analytics

The weekly analytics workflow uses Azure workload identity federation rather
than a client secret. A maintainer with permission to create Entra applications
and role assignments runs:

```powershell
.\infrastructure\azure\configure-analytics-oidc.ps1
```

The script grants the workflow only `Log Analytics Reader` on
`excelmcp-logs` and configures the repository's non-secret Azure identifiers.
It cannot write telemetry or change Azure resources after setup.

## Excel development VM

The VM template adapts the existing
[mcp-windows desktop runner](https://github.com/sbroenne/mcp-windows/tree/a76b3eeb075419986d0b00721bae5cdc896c2f81/infrastructure/azure).
It uses the same `Standard_D2s_v7` size and 128 GB Standard SSD, in a separate
resource group and network. Eligible Windows client development/test rights
are required. Excel installation, activation and unattended-use rights must
be verified separately.

After approving Azure spending, an agent or maintainer can provision it with:

```powershell
& .\scripts\Deploy-ExcelAgentRunner.ps1 -WhatIf
& .\scripts\Deploy-ExcelAgentRunner.ps1
```

`-WhatIf` does not contact Azure or change resources. Actual deployment uses
the signed-in Azure subscription, pins the selected Windows image version,
generates a password in a restricted temporary directory and stores it in
Key Vault. Reruns reuse the stored password and existing image. The helper
refuses unrelated resource groups and implicit VM edition/size changes.
Existing VMs must already be deallocated; provisioning refuses a running VM
instead of interrupting development work.
Temporary parameter files are removed even on failure.

Provisioning sets shutdown protection and deallocates the VM when successful.
If deployment fails, inspect the partial deployment and billing before rerunning.
The helper does **not** install Excel, register a runner, enable automatic
maintenance or change cloud-agent routing. Those steps remain required before
this machine can accept development tasks.

Public inbound access, including RDP, is blocked. Outbound internet access
allows HTTP and HTTPS, as required by Office CDN downloads and update/certificate
endpoints, but this is **not** a destination allowlist or a
replacement for GitHub's agent firewall. The reused persistent machine must
not run arbitrary untrusted contributor code. No additional Linux VM or
managed Azure Firewall is provisioned.

Deployment safety checks run without contacting Azure:

```powershell
& .\scripts\tests\deploy-excel-runner.tests.ps1
```

### Retail Excel installation and desktop preparation

After verifying the entitlement is specifically `Excel2024Retail`, use:

```powershell
& .\scripts\Install-ExcelAgentOffice.ps1
& .\scripts\Initialize-ExcelAgentDesktop.ps1
```

Run these sequentially, starting with the VM deallocated. Both commands leave
it deallocated afterwards. They do not register a coding runner or change
the repository's hosted setup.

The installer uses Microsoft's Office Deployment Tool and CDN, with a valid
Microsoft executable signature required before execution. It verifies the
installed retail product, x64 architecture, version and Excel executable.
Installation runs as a bounded SYSTEM task, outside an interactive coding job.
An existing matching installation is verified rather than replaced.
Installer timeouts and missing results are failures, not successful setup.

This is retail Excel, not Office LTSC. Its configured update channel is
`Current`, with Office automatic updates initially enabled. The maintenance
worker disables automatic Office servicing so updates run only in the protected
maintenance window. Windows Update alone does
not update Click-to-Run Excel. Activation is separate; an installed product
does not establish licence rights or successful activation.

Desktop preparation adapts `mcp-windows` Sysinternals Autologon and logon-task
setup. The `excelrunner` account is not a local administrator. Its generated
password is stored in Key Vault; temporary VM identity access is removed
before the desktop is restarted and qualified. An existing system identity
added by subscription policy is preserved, provided it has no direct Azure role
assignments or access to the dedicated password vault. Only an identity created
by the helper is removed. Cleanup has its own time budget and is verified before
continuing. Startup scripts are protected
against modification by that account. Profile setup enables VBA project access
for development tests, without disabling other macro protections.

The desktop check requires that account's Explorer session and a successful
profile task during the current boot. It does not yet establish Excel
activation or COM behavior. Actual workbook and CLI/MCP acceptance checks remain
required before the runner can accept work.

```powershell
& .\scripts\tests\install-excel-runner-office.tests.ps1
& .\scripts\tests\setup-excel-desktop.tests.ps1
& .\scripts\tests\runner-identity.tests.ps1
& .\scripts\tests\open-excel-activation.tests.ps1
```

### One-time private activation access

`excel-desktop-access.bicep` offers the free **Developer** Bastion SKU for
browser-based activation in supported regions. It never falls back to a
paid SKU. Its RDP rule permits only virtual-network traffic, not public
internet access. The desktop script's explicit `Activation` action grants
the non-admin account remote-desktop access without granting administrator
membership. Remove the temporary `AllowPrivateActivation` rule after activation
and restart to restore the automatic console session before running Excel tests.

This access is for the user's one-time sign-in/activation only; it does not
replace automated setup or qualify the machine for coding tasks.

`scripts\Open-ExcelAgentActivation.ps1` starts the prepared desktop, checks its
current-boot state and enables the desktop account's private remote access.
The private network rule from the template must already be present; the helper
does not recreate it after it has been closed. It leaves the VM running,
with automatic shutdown set one hour after the command starts.
`-CopyPasswordToClipboard` optionally copies the Windows password without
printing it; obtain the user's permission before replacing their clipboard.
It does not install or activate Excel. Activation remains a user sign-in step.

### Development toolchain

Run `scripts\Install-ExcelAgentToolchain.ps1` after desktop preparation, with the
VM deallocated. It follows the upstream prerequisite installer and SYSTEM worker
pattern and the SDK version and roll-forward policy in `global.json`. It installs
Git for Windows, PowerShell 7, Node.js 22 and signed 64-bit Python 3.13 with pip,
and verifies their actual executables and versions.
Downloads require a valid signature from the expected publisher. Installer
processes have bounded deadlines and retain their real exit codes. The worker
runs from an administrator-protected directory. Reruns retain matching tools
and refuse an unexpected Node.js major version rather than silently replacing it.

The helper refuses an unowned or running VM and a registered coding runner,
checks the policy-owned identity's direct roles and dedicated-vault access,
and deallocates the machine on success or failure. It does not enable a runner,
install evaluation dependencies, or change hosted CI or cloud-agent routing.
Activation, account cleanup and the full development acceptance test remain
separate gates.

With the current `latestFeature` policy it installs the latest stable SDK in the
requested major/minor channel, rather than assuming the minimum version has the
compiler needed by the source generators. The SDK resolver reads the actual
repository `global.json` from the protected provisioning directory, and the report
includes both the requested minimum and the resolved SDK. A `disable` policy
requires the exact version; unsupported policies fail explicitly.

### Whole-job control and automatic maintenance

The control design follows `mcp-windows`: GitHub-hosted workflows start and stop
one reused Windows VM. There is no second controller VM. GitHub controls Azure
through short-lived workload identity, not a stored Azure password. Control
permissions are limited to the dedicated resource group, except read-only
subscription role-assignment metadata needed to reject a privileged VM identity.
Azure shutdown scheduling also checks write permission on its linked VM, so
the dedicated-group role includes VM settings write access. It grants neither
role-assignment writes nor password-vault data access.
The coding desktop receives neither the control identity nor a GitHub
administration token.

Automatic issue-based development starts when the repository owner assigns an
open issue to the platform Copilot bot. The default-branch hosted controller
waits up to five minutes for approved runner demand, then keeps its shared
control slot through the exact admitted job and owned cleanup. It qualifies
the idle desktop between successive approved jobs, drains queued demand within
the same bounded control deadline, and parks the VM without needing a bot-triggered completion
workflow. Other actors, assignees and pull-request assignment events cannot
use this owner-only path. Existing coding-job identity and admission checks
still apply; assigning an issue does not authorize arbitrary runner work.
GitHub's workflow-approval policy is unchanged.

Standalone cloud Task/API starts do not emit the owner issue-assignment event.
They retain scheduled discovery and operator dispatch; scheduled GitHub Actions
can be delayed and are not a reliable immediate-start guarantee. Use owner
issue assignment for the automatic start/work/shutdown route.

Operator commands, run by an approved agent or maintainer rather than assigned
to the end user:

```powershell
& .\scripts\Invoke-ExcelRunnerMaintenance.ps1
& .\infrastructure\azure\configure-excel-control-oidc.ps1
& .\scripts\Register-ExcelAgentRunner.ps1
```

Run VM operations sequentially from a deallocated VM. Registration verifies the
official Windows x64 runner archive's published SHA256. A one-hour registration
token is encrypted to a transient guest key; its private key is protected by
the dedicated account's Windows data protection. Transient transport files
are removed. The repository runner is named `azure-excel-copilot`, has only the
`excel-copilot` custom label, and remains offline after registration. It runs
in the limited interactive desktop, never as a SYSTEM service.

The manually started runner task has **no logon trigger**. Each listener uses
`run.cmd --once`, which waits for a complete job and then exits while retaining
registration. Protected admission identifies one approved cloud-agent or
owner-dispatched validation run, its job, the current boot and a ten-minute
start deadline. The job-start hook refuses a different job before its setup
or code executes. Setup completion does not authorize shutdown.
Control rechecks the exact GitHub job after desktop preparation, before starting
a listener. Work already cancelled or completed is not admitted; an idle VM is
parked rather than left waiting for a job that no longer exists.
After admission, a short job or cancellation can finish before the first
listener check. Control rechecks the same job identity and completion instead
of reporting a false startup timeout. It still leaves active guest cleanup
undisturbed and requires idle recovery and desktop qualification before parking.

The hosted control workflow checks complete GitHub jobs and guest listeners,
workers, workbooks and cleanup records. It keeps active work undisturbed,
recovers cancelled jobs through the limited account, checks desktop recovery
and then deallocates the VM. Process cleanup uses retained handles, PID/start-time
identity and account ownership. Cleanup establishes account
ownership before accessing a process handle or start time: unrelated system
processes may be visible in the desktop session but unreadable by the limited
account. Unknown ownership still quarantines the runner, and stopping an owned
process still requires its exact PID/start-time identity. Workspace cleanup removes all contents of the
exact runner checkout and rejects directory links. It retains the empty checkout
directory because the completing runner can still hold it as its working
directory. An expired idle listener may be
drained only when no active whole job or worker is found.

Weekly maintenance installs approved Windows security, critical, definition
and update-rollup updates through bounded SYSTEM tasks. It repeats clean scans
after restart, separately services retail Excel through Click-to-Run, restarts
again and verifies the actual desktop. Excel must be x64 `Excel2024Retail`,
have a licensed perpetual device entitlement and calculate and save/reopen a
formula returning 42. A newer installed Current-channel Office release is
accepted; lagging release metadata must not cause a downgrade.

Patch evidence is protected and expires after eight days. Missing, partial,
failed or overdue checks prevent listener admission. Daily health creates
one open maintenance alert and schedules at most one maintenance retry a day
when no maintenance run is already active. Hosted maintenance history also
blocks repeated VM wakes for stale queued work. Control failures use a
deduplicated alert. Shared queued workflow concurrency keeps pending
maintenance from being replaced by frequent control checks.

The usual hosted control deadline is 25 minutes. Owner issue-assignment control
has an 85-minute deadline and a 110-minute workflow limit, including room for
the separate 15-minute failure-cleanup budget. Its demand wait is bounded to
five minutes, and active work is never reported as successful on timeout.
The generated Copilot job uses the supported 59-minute timeout so repository
setup does not consume a short development session's entire budget.
The six-hour coding shutdown backstop is refreshed before each admitted job.
Maintenance has a 175-minute
operation budget and a separate 15-minute cleanup budget. Guest update tasks
and host polling are bounded independently. Shutdown schedules provide
four-hour maintenance and six-hour coding backstops; these are hard protection
limits, not evidence of clean job completion. Failures preserve quarantine
and attempt deallocation only when complete-job and guest-idle checks allow it.

### Enablement and acceptance

All new workflows are opt-in. The **Actions** variable
`EXCEL_RUNNER_ENABLED=true` enables hosted control, maintenance, daily health
and owner-only validation. The separate repository **Agents** variable
`EXCEL_COPILOT_ENABLED=true` selects `excel-copilot` for cloud-agent setup.
Copilot does not receive Actions variables. Keep the **Actions** variable
`EXCEL_COPILOT_ENABLED=false` so ordinary setup workflow events continue to use
`windows-latest`; the same name belongs to two distinct variable stores.
The setup runner expression uses only the applicable variable, not generated
workflow display names or event contexts. Without the Agents opt-in, cloud
setup also retains the existing hosted Windows runner.

Copilot code review uses its separate `copilot-code-review.yml` setup on
`windows-latest`, regardless of the Agents switch. Without that file GitHub
reuses the cloud-agent setup, which would direct reviews to the Excel label.
Reviews are not admitted by the coding-job controller; keep their existing
hosted tooling separate rather than weakening the trusted cloud-job policy.
The review setup does not install or run the on-demand LLM evaluations.
Rejected guest admission reports only the GitHub run, attempt, job, actor
and event fields, never the full environment or personal desktop identity.
The coding job's runtime `GITHUB_ACTOR` is `copilot-swe-agent[bot]`, not the
REST API's `Copilot` display login. Hosted admission still verifies the
platform bot ID, dynamic workflow path and repository, then binds the exact
run/job and current boot; the guest must match the actual runtime bot identity.
The protected toolchain includes Git Bash, checksum-pinned Windows x64
jq and signed 64-bit Python 3.13, with their protected directories first on the
machine PATH. Python and pip are installed administratively before registration;
the limited coding account must not run `actions/setup-python`'s first-time
all-users installation. Hosted setup retains that action.
Readiness verifies Python's version, architecture, pip and protected PATH.
The native architecture probe works in both PowerShell versions. PATH checks
require the protected application first; a later Windows Python alias is not
an override.
Repository, extension and shared npm setup use
`scripts\Invoke-CopilotSetupNpm.ps1`. On this desktop it calls the protected
Node.js installation's `npm.cmd`, retaining each step's working directory.
GitHub's downloaded agent runtime can prepend its own incomplete npm shim
to PATH; that shim must not replace the provisioned dependency installer.
Hosted setup retains the npm selected by `actions/setup-node`. A missing
protected command or failed installation stops setup explicitly.
Complete development-tool provisioning before registration; the regular
Windows/Excel maintenance worker does not install or update these tools.
GitHub's generated
initialization uses `bash` and `jq` before repository setup steps, even on
Windows; adding them in `copilot-setup-steps.yml` is too late. Qualification
must resolve the protected binaries and execute jq through Git Bash before
admitting a listener. Existing registered desktops can run the same cloud
prerequisite helper only during guarded, idle administrative maintenance.
The same helper enables Windows Developer Mode before admission. GitHub's
generated Windows runtime archive contains symbolic links; native `tar`
otherwise fails to extract them under the limited desktop identity before
repository setup. This permits supported unprivileged link creation without
adding the coding account to Administrators or enabling Device Portal.
Toolchain readiness rejects missing or disabled Developer Mode. Qualification
must also establish native archive-link extraction in the actual non-admin
desktop, not only an administrator's successful extraction.

Operators can read the actual firewall state through
`GET /repos/{owner}/{repo}/copilot/cloud-agent/configuration` and manage the
Agents switch through `POST /repos/{owner}/{repo}/agents/variables` or
`PATCH /repos/{owner}/{repo}/agents/variables/EXCEL_COPILOT_ENABLED`.
Use the documented `X-GitHub-Api-Version: 2026-03-10` header. Confirm
`is_firewall_enabled=false` before Windows routing; an old Actions firewall
variable is not evidence of the current Copilot setting.

The `excel-runner-control` environment must allow only the default branch.
Bootstrap refuses an existing broader environment rather than overwriting
reviewer, wait or branch protections. Azure identifiers are non-secret
environment variables. Registration and federation do not enable either
opt-in variable.

The workflows must be published on the default branch before scheduled
operation and the Copilot setup can be proven. Keep routing disabled until a
real registered-runner validation job proves whole-job ownership, queue and
cancellation recovery, workbook state, both entry points and shutdown.
Then run a controlled cloud-agent pilot before broad use. Installing Office,
passing source safety tests or completing a local desktop check is not proof
that GitHub's generated agent job uses this runner.

This persistent VM is for trusted, maintainer-approved development. Cleanup
does not make a reused desktop equivalent to a disposable security boundary.
HTTP/HTTPS outbound access is not a destination allowlist. Retail activation
does not itself establish unattended-use or Windows development/test rights.
No personal account sign-in, activation reset or workbook copying is part
of normal maintenance.

Run its offline checks with:

```powershell
& .\scripts\tests\excel-runner-toolchain.tests.ps1
& .\scripts\tests\excel-runner-toolchain-host.tests.ps1
& .\scripts\tests\copilot-setup-npm.tests.ps1
```

### Qualification and account privacy

Keep the coding runner unregistered until the desktop, account cleanup, Excel
licence and full development checks have been qualified. A successful installer
or formula calculation alone is not a complete readiness check. Use the selected
repository SDK to build Release with zero warnings, then run
`scripts\Test-E2E.ps1` sequentially in the non-admin interactive desktop.
All required CLI, rebuild and MCP stages must actually run and pass.

One-time personal-account activation requires explicit user approval. Sign-out
must not be treated as proof that Windows, OneDrive, browser or broker sign-ins
were removed. Obtain separate approval before removing local personal files,
never delete cloud files as part of that cleanup, and never sign out other
devices. Close the temporary private activation network rule afterwards.

Windows can recreate device credentials after sign-out. Distinguish exact
device targets from personal-account targets using non-secret metadata; do not
read credential blobs, print account names or blindly delete returning caches.
The presence of a generic cache file alone does not prove a reusable personal
sign-in. Conversely, zero counts from an unavailable account-inspection method
are not successful verification.

Use edition-appropriate activation evidence. Microsoft's published `vNextDiag`
guide covers Microsoft 365, while `OSPP.VBS` supports volume activation. Empty
reports from those tools do not by themselves prove that Excel 2024 Retail
activation failed. Do not reset Office licence files to make an inspection pass,
switch to LTSC without a matching entitlement, or retain a personal sign-in to
avoid investigating the installed edition.
