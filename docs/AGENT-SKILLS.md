# Agent guidance

ExcelMcp does not ship agent skills. The CLI plugin contains its argument-safe
launcher; the MCP plugin contains its server connection configuration. The VS Code
extension bundles the MCP runtime without a chat-skill contribution.

## Discover supported operations

For the CLI, use `excelcli --help`, `excelcli list-actions <command>`, then
`excelcli <command> <action> --help` for the actions needed. The plugin launcher
is `bin\start-cli.ps1` inside the installed CLI plugin; it runs
`npx -y @sbroenne/excelcli@latest` without adding `excelcli` to PATH.

For MCP, use the tool descriptions and input schemas. Both entry points expose
the same workbook operations. Native help and schemas define supported inputs,
defaults, and results; do not infer syntax from general Excel knowledge.

## Optional documentation

The [reference guides](reference/README.md) cover workflows and limitations.
The [report-formatting guide](reference/report-formatting.md) remains available
when a user requests presentation advice. It is not automatically loaded or
packaged as a skill. Match the user's instructions and existing template first.

## Authoring and packaging

| Content | Source |
|---------|--------|
| Operation and parameter descriptions | Core interface XML documentation and attributes |
| Generated action metadata | Shared generators and command contracts |
| Minimal MCP instructions | MCP Server `Program.cs` |
| Workflows and formatting advice | `docs/reference/` |
| Plugin launchers and configuration | `.github/plugins/` |

Keep product-specific warnings beside the affected action. Do not copy generic
Excel tutorials or action catalogs into automatically loaded instructions.
Follow [product guidance rules](agents/product-guidance.md) when changing metadata.

Build plugins with `scripts\Build-Plugins.ps1 -Version <version>`.
`scripts\Build-ReleasePackages.ps1` owns complete distributable preparation.
Neither command prepares or installs skills. Packaging does not uninstall older
product copies from user machines, and source changes do not publish plugins.
