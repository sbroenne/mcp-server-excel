# Run VBA Macros from an AI Agent

Many Excel workbooks carry decades of VBA. ExcelMcp lets an AI assistant read,
write, and **execute** that code inside the real Excel application — so existing
macros keep working instead of being rewritten.

This is something file-parser libraries cannot do at all: `.xlsm` macro code is
only meaningful to Excel's VBA host.

## One-time setup for project access: enable VBA trust

Excel blocks all programmatic access to the VBA project by default. **You must
enable it manually** before listing, viewing, importing, updating, or deleting
VBA modules. ExcelMcp never changes this setting for you, because doing so
silently would be a security problem. Running an existing macro does not inspect
the VBA project and does not require this setting.

1. Open Excel
2. **File → Options → Trust Center → Trust Center Settings**
3. **Macro Settings**
4. Tick **Trust access to the VBA project object model**
5. Click OK, then restart Excel

Without it, VBA project inspection and editing fail with a `Permissions` error.
`vba status` instead reports blocked project access and the required manual step.
This is a per-machine, per-Office-install setting, so remote and CI machines that
manage VBA modules need it too.

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
excelcli -q vba status --session $session
excelcli -q vba references --session $session
excelcli -q vba view --session $session --module-name Module1
```

`list` returns every module, class module, form, and document module. `view`
returns the full source of one module.

`status` reports actual project access, password protection (`None` or `Locked`),
and execution mode (`Design`, `Run`, or `Break`). If access is blocked,
`projectAccess` is false and `accessMessage` explains why; protection and mode
are not guessed. A locked project must be unlocked manually in Excel. These
checks never enable trust, unlock a project, or stop a running macro.

`references` lists the project's library references and reports
`hasBrokenReferences`. Each entry has an `index` and `isBroken`. Healthy entries
also include `name`, `description`, `libraryId`, version numbers, and `builtIn`.
Broken-reference metadata is unavailable and returned as null because Excel
can fail when reading it. No libraries are added, removed, or repaired, and
local library paths are not returned.

Find source without returning entire modules:

```powershell
excelcli -q vba search --session $session --search-text GenerateReport
excelcli -q vba search --session $session --module-name Module1 --search-text customerId --whole-word --match-case --max-matches 20
```

Search is literal text, not a wildcard or regular expression. It includes
comments and string literals. The default is a case-insensitive search across
all modules, returning up to 50 matches; `maxMatches` can be 1 through 100.
Each match contains `moduleName`, one-based `line` and `column`, and an
`excerpt` of at most 200 characters. `hasMore` means additional matches were
omitted; narrow the text or select a module to inspect them.

For MCP, the search inputs are `search_text`, `module_name`, `whole_word`,
`match_case`, and `max_matches`; the CLI flags above use hyphens.
Use the read-only MCP tool `vba_read` with actions `search`, `references`, and `status`.

For smaller reads, `list` also reports each procedure's kind and source lines.
`startLine` includes introductory comments and blank lines; `bodyStartLine`
identifies the actual `Sub`, `Function`, or `Property` declaration.
Read one procedure or a line range of up to 500 lines:

```powershell
excelcli -q vba read --session $session --module-name Module1 --procedure-name GenerateReport --procedure-kind Sub
excelcli -q vba read --session $session --module-name Module1 --start-line 1 --line-count 40
```

The result includes `startLine`, `returnedLineCount`, `hasMore`, and
`nextStartLine`. When a procedure is longer than 500 lines, use `nextStartLine`
to read the next part. A procedure read also returns `sourceHash`; keep it to
guard a later edit against changes made since the read.

## Run a macro

```powershell
excelcli -q vba run --session $session --procedure-name "Module1.GenerateReport" --timeout-seconds 120
```

The procedure name uses `Module.Procedure` form. Pass arguments with
`--parameters` when the macro takes them.

Set a timeout that matches the work. A macro that waits on a dialog will otherwise
hold the session until the default limit expires.

If the requested run timeout expires, the operation is reported as a timeout
rather than a cancellation. If execution has started, Excel may still be busy
running the macro, so the workbook session is closed; open the workbook again
before continuing. If the timeout expires while the operation is still queued,
the macro is not started and the session remains open. Follow the returned error
message to determine whether reopening is required.

## Add or update code

```powershell
excelcli -q vba update --session $session --module-name Module1 --vba-code $code
excelcli -q vba import --session $session --module-name Helpers --vba-code-file .\Helpers.bas
```

`update` replaces the whole module body. `import` adds a module from a `.bas`
file. `delete` removes a module.

To change one procedure without replacing the rest of the module, pass the
procedure's `sourceHash` from `vba read` to `replace-procedure`:

```powershell
excelcli -q vba replace-procedure --session $session --module-name Module1 --procedure-name GenerateReport --procedure-kind Sub --expected-source-hash $sourceHash --vba-code $updatedProcedure
```

The replacement must contain exactly one complete procedure with the same name
and kind. If the source changed since it was read, the edit is rejected; read it
again before deciding what to change. The result includes the stored source's
new `sourceHash`. Reading it back confirms only the saved text, not that the
procedure compiles or runs correctly.
Introductory comments and surrounding blank lines are preserved; supply only
the replacement declaration through its matching `End` statement. The source
fingerprint includes introductory comments, so intervening changes to those
comments also reject the replacement. Replacement is blocked while the project
is running or paused in break mode.

## Save to the right file format

VBA commands support `.xlsm` workbooks. Saving VBA into an `.xlsx` silently
discards it. If you are adding VBA to an `.xlsx`, save-as `.xlsm` first.

## Verify

After running a macro, check the effect rather than trusting a success flag:

```powershell
excelcli -q range get-values --session $session --sheet Summary --range A1:D20
excelcli -q screenshot capture-sheet --session $session --sheet Summary
```

## Known gotchas

**Access denied while listing or changing modules** can mean disabled project
trust or a password-locked project. Use `vba status` to distinguish them.
This is separate from a failure returned by `run`, which comes from Excel
while locating or executing the requested macro.

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
