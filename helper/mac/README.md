# ExcelMcp native Mac helper

**Status: Build/readiness infrastructure implemented; Excel artifact and
object-model acceptance pending.** The current source supplies only the
read-only `helper.info` primitive. It does not enable Power Query or any other
gated public action.

## Independent version

`VERSION` is the helper semantic version, independent of the server's release
version. Change it only when helper source changes. The server checks compatible
major and required primitive names, not equality with its own version.

`ExcelMcpHelper.bas.template` is reviewed VBA source. The preparation command
injects `VERSION` into the import module:

```powershell
pwsh -NoProfile -File scripts/Build-MacHelper.ps1 -PrepareOnly
```

This produces `artifacts/mac-helper/ExcelMcpHelper.bas` without opening Excel.
Do not edit that generated module.

## First maintainer bootstrap

Excel's Apple Events dictionary cannot author or import VBA modules. In desktop
Excel, create an ordinary macro-enabled workbook, manually import the prepared
module in the VBA editor, and save it as `NativeHelperBootstrap.xlsm` outside
the repository. Keep it open for the build. Approve its macros interactively;
the scripts do not change trust/security settings.

This manual import is a maintainer bootstrap, not a user installation step.
Never manufacture VBA binaries or workbook package internals.

## Local Excel-authored artifact

Build the Release CLI, then run:

```powershell
pwsh -NoProfile -File scripts/Build-MacHelper.ps1 `
    -BootstrapPath /absolute/path/NativeHelperBootstrap.xlsm
```

The CLI's internal Service maintenance route invokes native Apple Events.
Before building, it verifies exact bootstrap workbook identity and that its
handshake version matches the reviewed helper source. Excel saves the add-in
with its native `.xlam` format. The output includes `ExcelMcpHelper.xlam` and
`SHA256SUMS` under `artifacts/mac-helper/`.

The command refuses an existing output; it does not silently overwrite an
artifact. A timeout may leave Excel's build outcome uncertain. Reconcile the
bootstrap and artifact before retrying. Do not automatically delete, close, or
retry an uncertain build. Save As changes the live bootstrap workbook identity
to the add-in, without editing the original `.xlsm` file.

## Acceptance and independent release

Before publication, install the artifact using Excel's add-in manager and run:

```powershell
'{"command":"service.helper-check"}' | excelcli -q batch
```

Verify the installed version and primitives, an absent helper, incompatible
major, missing primitive, macro approval behavior, and real CLI/MCP acceptance
for every action intended to use it. A successful handshake is not proof that
Power Query or formula APIs preserve the public contract.

After that acceptance and separate publication authorization, publish only the
Excel-authored `.xlam` and checksum under a `helper-vX.Y.Z` GitHub release.
Do not tie that release to server builds or rebuild an unchanged helper during
ordinary server releases. Users download the artifact and follow
[the installation guide](../../docs/INSTALLATION.md#optional-macos-native-helper);
they do not import source modules.
