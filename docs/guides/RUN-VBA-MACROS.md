# Run VBA Macros from an AI Agent

Many Excel workbooks carry decades of VBA. ExcelMcp lets an AI assistant read,
write, and **execute** that code inside the real Excel application — so existing
macros keep working instead of being rewritten.

This is something file-parser libraries cannot do at all: `.xlsm` macro code is
only meaningful to Excel's VBA host.

!!! note "Platform availability"
    These commands currently run through the Windows COM backend. On macOS,
    ExcelMcp checks the existing Office macro and VBA project-model preferences
    without prompting or changing them, then returns an explicit capability
    error. Macro execution remains gated until a repository-owned `.xlsm`
    fixture proves unattended workbook-qualified execution through both CLI and
    MCP. The Mac distribution includes reviewable source for an optional,
    versioned helper add-in. Its fixed Power Query lifecycle and VBA
    project-model implementations remain gated until each method passes real
    CLI and MCP evidence with the required user-managed trust.

## Optional macOS helper setup (preview)

The packaged `helpers/ExcelMcpHelper.bas` is repository-owned source for a fixed,
allowlisted dispatcher. It is not installed automatically, and ExcelMcp never
imports it into a user workbook or changes either VBA security setting.

To prepare the helper for later capability probes:

1. Review `helpers/ExcelMcpHelper.bas` from the same ExcelMcp build you installed.
2. In Excel for Mac, create a new blank workbook and open the Visual Basic
   Editor.
3. Import the reviewed `.bas` file as a standard module.
4. In Excel, save that new workbook as an Excel add-in named exactly
   `ExcelMcpHelper.xlam`. Do not manufacture or replace `vbaProject.bin`.
5. Enable that exact add-in through Excel's add-in manager.
6. Set `EXCELMCP_MAC_VBA_HELPER_PATH` for the process that starts ExcelMcp to
   the add-in's exact absolute path.

The configured file name, open add-in `FullName`, helper version, protocol
version, request correlation, target workbook `FullName`, action, and argument
shape are all checked before a helper operation. Requests and responses are
bounded to 262,144 UTF-8 bytes and the dispatcher has no arbitrary evaluation
of VBA or AppleScript. Its run action accepts only a validated
`Module.Procedure` identity in the exact target workbook and at most 30 string
parameters. The gated Power Query actions do accept M source for query
authoring and temporary evaluation. Installation
alone does not enable a command: production actions stay gated until their
individual methods have real-Excel evidence.

Helper version `1.4.0` adds fixed Power Query create/update/refresh/refresh-all,
load-to/unload, and temporary-query evaluation candidates. They support only
connection-only or one exact worksheet-table destination; Data Model and
multi-destination variants fail explicitly. The source implements rollback and
temporary-artifact cleanup, but every corresponding proof flag remains false
until prompt-free real-Excel CLI and MCP acceptance succeeds.

Helper 1.4.0 also has source-complete candidate routes for VBA `list`, `view`,
`import`, `update`, and `delete`, plus exact workbook-qualified `run`. They
remain unavailable by default. Maintainers may enable only named actions with
`EXCELMCP_MAC_VBA_CANDIDATE_ACTIONS` while running the guarded public
acceptance workflow. This opt-in collects evidence; helper presence and
permissive trust settings never enable actions on their own.

The guarded workflow is `scripts/Test-MacVbaPublicAcceptance.ps1`. It requires
an existing Excel-authored `.xlsm` containing a repository-owned marker
procedure that writes its one string argument to a dedicated cell. It performs
source lifecycle, marker execution, and public range verification through both
CLI and MCP, then closes without saving. Do not treat `-ValidateOnly` as Excel
evidence.

Maintainers running bounded candidate acceptance may set
`EXCELMCP_MAC_POWERQUERY_CANDIDATE_ACTIONS` to a comma-separated list of exact
actions such as `powerquery.create,powerquery.delete`. This opt-in enables only
listed actions that the matching helper version also advertises; unknown names
are rejected, and it never enables another Power Query method implicitly.
Remove the variable after the acceptance run.

The same version adds a read-only `helper.inspect-engines` probe for the
late-bound `Workbook.XmlMaps` and `Workbook.Model` object paths. It reports only
the observation made against the exact target workbook: `accessible`,
`unavailable`, `error`, or `unknown`, plus a fixed reason code and optional
object count. A zero-table workbook model is `unknown`, not evidence that the
engine is unavailable. An accessible XML Maps collection proves only object
model access, not XML import, export, schema, or XPath behavior.

The first installed-helper validation must also exercise the VBA parser itself,
not only the host DTOs: a protocol version such as `1.4`, malformed JSON,
unknown/duplicate properties, and mismatched correlation must be rejected
before any mutation, followed by one read-only capability request through both
CLI and MCP.

After direct helper-engine validation, run the separately guarded public Power
Query lifecycle workflow only with an exclusive desktop Excel slot:

```powershell
pwsh ./scripts/Test-MacPowerQueryPublicAcceptance.ps1 `
  -HelperPath '/absolute/path/ExcelMcpHelper.xlam' `
  -WorkbookPath '/absolute/path/ExcelMcpPowerQueryAcceptance.xlsx' `
  -HelperInstalledTrustedConfirmed `
  -ExcelAuthoredWorkbookConfirmed `
  -DedicatedWorkbookConfirmed `
  -ExcelSlotConfirmed
```

The workbook must be an Excel-authored dedicated workbook with no existing
Power Queries. The workflow uses only public CLI and MCP commands, literal
credential-free `#table` M, separate disposable working copies, and an exact
per-action candidate allowlist scoped to child processes. It verifies supported
worksheet and connection-only lifecycle behavior, exact loaded `A1:A2` values
through public range reads after saved reopen checkpoints, plus continued
rejection of Data Model and combined destinations. `-ValidateOnly` checks the
plan and emits a non-proof receipt; it does not launch Excel or prove a public
method.
The CLI working copy uses a short unique daemon pipe. Cleanup invokes bounded
`service stop` on that same pipe, verifies its result, and never targets shared
Excel or another client's daemon.

Macro execution and VBA project access are separate settings. Do not enable all
macros globally to install the helper. Enable only the trust your reviewed
workflow requires. To remove the helper, disable it in Excel's add-in manager,
close only that exact add-in if it is open, delete the `.xlam` if desired, and
remove `EXCELMCP_MAC_VBA_HELPER_PATH`.

## One-time setup: enable VBA trust

Excel blocks all programmatic access to the VBA project by default. **You must
enable it manually** — ExcelMcp never changes this setting for you, because doing
so silently would be a security problem.

The setting enables source inspection and mutation; it is not required merely
to invoke an already trusted macro. Macro execution is governed separately by
Excel's macro security and per-workbook trust.

On Windows, use **File → Options → Trust Center → Trust Center Settings →
Macro Settings**. On Mac, use Excel's corresponding **Security** preferences.
Enable **Trust access to the VBA project object model**, then restart Excel.

Without it, every VBA operation fails with an access error. This is a per-machine,
per-Office-install setting, so remote and CI machines need it too.

!!! warning "Security implication"
    Enabling VBA trust allows any program on the machine to read and modify VBA
    code in workbooks you open. Enable it only if you actually need VBA
    automation, and only on machines you control.

## What you ask for

> List the macros in `report.xlsm` and show me what `Module1` does.

> Run the `GenerateReport` macro in `monthly.xlsm` and save the result.

## Inspect before you run

```powershell
$session = (excelcli -q session open C:\books\report.xlsm | ConvertFrom-Json).sessionId

excelcli -q vba list --session $session
excelcli -q vba view --session $session --module-name Module1
```

`list` returns every module, class module, form, and document module. `view`
returns the full source of one module.

## Run a macro

```powershell
excelcli -q vba run --session $session --procedure-name "Module1.GenerateReport" --timeout 120
```

The procedure name uses `Module.Procedure` form. Pass arguments with
`--parameters` when the macro takes them.

Set a timeout that matches the work. A macro that waits on a dialog will otherwise
hold the session until the default limit expires.

## Add or update code

```powershell
excelcli -q vba update --session $session --module-name Module1 --vba-code $code
excelcli -q vba import --session $session --module-name Helpers --vba-code-file .\Helpers.bas
```

`update` replaces the whole module body. `import` adds a module from a `.bas`
file. `delete` removes a module.

## Save to the right file format

Macro-enabled workbooks must be `.xlsm` (or `.xlsb`). Saving VBA into an `.xlsx`
silently discards it. If you are adding VBA to an `.xlsx`, save-as `.xlsm` first.

## Verify

After running a macro, check the effect rather than trusting a success flag:

```powershell
excelcli -q range get-values --session $session --sheet Summary --range A1:D20
excelcli -q screenshot capture-sheet --session $session --sheet Summary
```

## Known gotchas

**Access denied on every VBA action** means the trust setting above is off. It is
by far the most common cause of VBA failures.

**Macros can display dialogs.** A `MsgBox` inside a macro blocks execution until
someone dismisses it. Prefer macros that write results to cells over ones that
prompt. Always pass a timeout.

**Macros run with full user privileges.** A macro can touch the file system,
network, and other applications. Review code before running it, especially code an
assistant generated or a workbook you did not author.

**Line continuations and quoting.** When passing VBA source on a command line,
prefer `import` from a `.bas` file — it avoids shell-escaping problems entirely.

**Excel must be installed.** VBA execution is not emulated; it runs in Excel's own
VBA host.

## Related

- [Advanced automation operations](../features/AUTOMATION-ADVANCED.md)
- [CLI installation and setup](../INSTALLATION-CLI.md)
- [Security policy](../../SECURITY.md)
