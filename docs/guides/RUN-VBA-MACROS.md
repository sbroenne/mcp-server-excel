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
    MCP. Source list/view/import/update/delete remain gated because Excel's
    installed Apple Events dictionary exposes no VBA project-model route.

## One-time setup: enable VBA trust

Excel blocks all programmatic access to the VBA project by default. **You must
enable it manually** — ExcelMcp never changes this setting for you, because doing
so silently would be a security problem.

The setting enables source inspection and mutation; it is not required merely
to invoke an already trusted macro. Macro execution is governed separately by
Excel's macro security and per-workbook trust.

1. Open Excel
2. **File → Options → Trust Center → Trust Center Settings**
3. **Macro Settings**
4. Tick **Trust access to the VBA project object model**
5. Click OK, then restart Excel

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
