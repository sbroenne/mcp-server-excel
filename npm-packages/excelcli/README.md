# ExcelMcp CLI

Run the self-contained Excel automation CLI through npm:

```powershell
npx -y @sbroenne/excelcli --help
npx -y @sbroenne/excelcli -q session open "C:\Data\Book.xlsx"
```

For repeated use, install the `excelcli` command on your PATH:

```powershell
npm install --global @sbroenne/excelcli
excelcli --version
```

Requires Node.js 18 or later and either Windows x64/Arm64 (x64 emulation) with
Microsoft Excel 2016+, or Apple Silicon macOS with Excel for Mac 16.112+.
No separate .NET runtime is needed. Keep
optional dependencies enabled so npm installs the matching native runtime.

The launcher forwards arguments, standard input/output, and exit codes to the
existing CLI. Excel operations and session management are unchanged.

[Documentation](https://excelmcpserver.dev/installation-cli/) |
[Source](https://github.com/sbroenne/mcp-server-excel) |
[Issues](https://github.com/sbroenne/mcp-server-excel/issues)
