# Installing the MCP Server - ExcelMcp

Installation instructions for the ExcelMcp **MCP Server** — the entry point for AI assistants like GitHub Copilot, Claude Desktop, Cursor, and any other MCP client. Looking for the CLI instead? See the [CLI Installation Guide](INSTALLATION-CLI.md).

## System Requirements

### Required
- **Windows OS** (Windows 10 or later)
- **Microsoft Excel 2016 or later** (Desktop version - Office 365, Professional Plus, or Standalone)
- **An interactive Windows desktop** with Excel available to the signed-in user

| Installation method | Additional requirements |
|---|---|
| VS Code extension | VS Code with GitHub Copilot; no separate Node.js or .NET |
| Claude Desktop MCPB | Claude Desktop and Node.js 18+ with npm/npx available on PATH; no separate .NET |
| npm / Copilot plugin | Node.js 18+ with npm/npx; no separate .NET |
| Standalone ZIP | No Node.js or .NET |
| NuGet tool | .NET 10 Runtime or SDK |

npm-based launches need network access to download packages and check for
updates. They use normal npm resolution and caching, not a guaranteed online
check on every launch. Install the current Node.js LTS for npm or MCPB setup:
`winget install OpenJS.NodeJS.LTS`.

### Windows Architecture

The VS Code extension bundles a server matching its package: x64 for Windows
x64 VS Code and native ARM64 for Windows ARM64 VS Code. No separate Node.js
installation is needed.

With npm, including the npx-based Claude Desktop MCPB, ARM64 Node.js selects
the native ARM64 server; x64 Node.js selects the x64 server, including on
ARM64 Windows. Standalone ZIP downloads currently bundle the x64 server.

### Optional (for specific features)
- **Microsoft Analysis Services OLE DB Provider (MSOLAP)** - Required for DAX query execution (`evaluate`, `execute-dmv` actions)
  - Easiest: Install [Power BI Desktop](https://www.microsoft.com/en-us/power-platform/products/power-bi/desktop) (includes MSOLAP)
  - Alternative: [Microsoft OLE DB Driver for Analysis Services](https://learn.microsoft.com/analysis-services/client-libraries)

---

## Quick Start (Recommended)

Use this order to avoid setup confusion:

1. **Choose one primary setup path**:
   - **VS Code Extension** (GitHub Copilot users) — bundles the server and Excel skill
   - **Claude Desktop MCPB** — one-click MCP installation
   - **GitHub Copilot Plugin** (Copilot CLI users) — marketplace installation
   - **npm package** (other MCP clients) — runs the self-contained server through `npx`
   - **Manual MCP setup** (other MCP clients like Cursor, Windsurf)
2. **Validate MCP setup** (run the quick test prompt in Step 4 of manual setup, or test in your client after extension/MCPB/plugin install)
3. **Optional:** also install the [CLI](INSTALLATION-CLI.md) (`excelcli`) for scripting/RPA

### VS Code Extension (Easiest - One-Click Setup)

1. **Install the Extension**
   - Open VS Code
   - Press `Ctrl+Shift+X` (Extensions)
   - Search for **"ExcelMcp"**
   - Click **Install**

2. **Open Copilot Chat**
   - Use a chat that supports tools.

3. **Ask Copilot to work with Excel**
   - Use a workbook path available on your Windows desktop.
   - Try: "Create an empty Excel file called test.xlsx."
   - With VS Code's default settings, the bundled **excel-mcp** server starts
     automatically when your request needs Excel tools. Approve server or
     tool use if prompted.
   - Copilot can load `excel-mcp-report-formatting` for requested report presentation.
     Type `/skills` to open VS Code's Configure Skills menu.

The extension includes a self-contained MCP server and its optional formatting skill.
No separate .NET, Node.js, CLI, or skill installation is needed. The CLI is
not included; install it separately if needed. Installing the extension does
not start an Excel workbook or approve server access for you.

**Marketplace Link:** [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)

---

### Claude Desktop (One-Click Install)

**Best for:** Claude Desktop users who want the simplest installation

Install Node.js LTS first (`winget install OpenJS.NodeJS.LTS`) if `npx` is not
already available. Restart Claude Desktop after changing PATH.

1. Download `excel-mcp-{version}.mcpb` from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Double-click the `.mcpb` file (or drag-and-drop onto Claude Desktop)
3. Restart Claude Desktop

The bundle configures Claude to run
`npx -y @sbroenne/mcp-server-excel@latest` directly. It contains no custom
launcher, npm installation, or fixed server executable. No separate .NET
installation is needed. The first launch downloads the Windows server; later
launches resolve `@latest` using npm's cache policy. Claude's built-in Node.js
does not guarantee that the external `npx` command is available.

**Already installed an older, binary MCPB?** Install the new npx-based bundle once.
Restarting an older bundle does not replace its fixed server executable.
The bundle itself still needs manual replacement for configuration changes;
fetching a newer npm server does not update the installed `.mcpb`.

---

### GitHub Copilot Plugin

**Best for:** GitHub Copilot CLI users who want plugin marketplace installation

```powershell
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install the MCP Server plugin
copilot plugin install excel-mcp@mcp-server-excel-plugins
```

**Note:** Plugin updates are published only when distributed plugin content changes.
The plugin version can lag the product release; its npx launcher uses the latest npm runtime.

---

## Manual MCP Setup (All MCP Clients)

**Best for:** Other MCP clients (Cursor, Windsurf, Cline, Claude Code, Codex), advanced users

### Step 1: Choose npm or the Standalone Executable

#### Option A: npm (Recommended)

Run the self-contained server directly through npm:

```powershell
npx -y @sbroenne/mcp-server-excel@latest --version
```

The npm package includes the Windows server, so it does not require .NET or a
separate download from GitHub Releases. npm caches the package after the first
run. `@latest` selects the release marked latest when starting the server;
it does not replace an already running server.

Windows x64 and ARM64 are supported. ARM64 Node.js selects
`@sbroenne/mcp-server-excel-win32-arm64`; x64 Node.js selects
`@sbroenne/mcp-server-excel-win32-x64`, including through emulation on ARM64
Windows. Keep optional dependencies enabled; do not use `--omit=optional`.
If the matching runtime is missing, the launcher reports an error rather than
falling back to another architecture.

#### Option B: Standalone Executable

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download **`ExcelMcp-MCP-Server-{version}-windows.zip`**
3. Extract the ZIP to a permanent location (e.g., `C:\Tools\ExcelMcp\`)

```powershell
# Example extraction
Expand-Archive "ExcelMcp-MCP-Server-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp"
```

The ZIP contains `mcp-excel.exe` — a fully self-contained executable (no .NET runtime needed).

### Step 2: Add the Standalone Executable to PATH

Skip this step when using npm.

To use `mcp-excel` as a command without specifying the full path:

```powershell
# Add to user PATH (persistent)
$toolsDir = "C:\Tools\ExcelMcp"
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*$toolsDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$userPath;$toolsDir", "User")
    Write-Host "Added $toolsDir to user PATH. Restart your terminal to apply."
}
```

Or manually: **Settings → System → About → Advanced system settings → Environment Variables → User variables → Path → Edit → New** → add `C:\Tools\ExcelMcp`

### Step 3: Configure Your MCP Client

#### Option A: Auto-Configure All Agents (Recommended)

Use [`add-mcp`](https://github.com/neondatabase/add-mcp) to configure all detected coding agents with a single command:

```powershell
npx add-mcp "npx -y @sbroenne/mcp-server-excel@latest" --name excel-mcp
```

This auto-detects and configures **Cursor, VS Code, Claude Code, Claude Desktop, Codex, Zed, Gemini CLI**, and more. Use flags to customize:

```powershell
# Configure specific agents only
npx add-mcp "npx -y @sbroenne/mcp-server-excel@latest" --name excel-mcp -a cursor -a claude-code

# Configure globally (user-wide, all projects)
npx add-mcp "npx -y @sbroenne/mcp-server-excel@latest" --name excel-mcp -g

# Non-interactive (skip prompts)
npx add-mcp "npx -y @sbroenne/mcp-server-excel@latest" --name excel-mcp --all -y
```

> **Requires:** [Node.js](https://nodejs.org/) for `npx`. Install with `winget install OpenJS.NodeJS.LTS` if not already available. No permanent `add-mcp` installation is needed.

> **Standalone alternative:** Use `npx add-mcp "C:\Tools\ExcelMcp\mcp-excel.exe" --name excel-mcp` after extracting the ZIP.

#### Option B: Manual Configuration

**Quick Start:** Ready-to-use config files for all clients are available in [`examples/mcp-configs/`](../examples/mcp-configs/)

**For GitHub Copilot (VS Code):**

Create `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"]
    }
  }
}
```

> For the standalone executable, replace `command` with `mcp-excel` and omit `args`.

**For GitHub Copilot (Visual Studio):**

Create `.mcp.json` in your solution directory or `%USERPROFILE%\.mcp.json`:

```json
{
  "servers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"]
    }
  }
}
```

**For Claude Desktop:**

1. Locate config file: `%APPDATA%\Claude\claude_desktop_config.json`
2. If file doesn't exist, create it with the content below
3. If file exists, merge the `excel-mcp` entry into your existing `mcpServers` section

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"],
      "env": {}
    }
  }
}
```

4. Save and restart Claude Desktop

**For Cursor:**

1. Open Cursor Settings (Ctrl+,)
2. Search for "MCP" in settings
3. Open its MCP configuration, or create `%USERPROFILE%\.cursor\mcp.json`
   for all projects (`.cursor\mcp.json` for one project)
4. Add this configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"],
      "env": {}
    }
  }
}
```

5. Save and restart Cursor

**For Cline (VS Code Extension):**

1. Install Cline extension in VS Code
2. Open Cline panel and click the MCP settings gear icon
3. Add this configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"],
      "env": {}
    }
  }
}
```

4. Save and restart VS Code

**For Windsurf:**

1. Open Windsurf Settings
2. Use **Open MCP config file** in the client's MCP settings
3. Add this configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"],
      "env": {}
    }
  }
}
```

4. Save and restart Windsurf

### Step 4: Validate MCP Setup

Restart your MCP client, then ask:
```
Create an empty Excel file called "test.xlsx"
```

If it works, you're all set! 🎉

**💡 Tip:** Want to watch the AI work? Ask:
```
Show me Excel while you work on test.xlsx
```
This opens Excel visibly so you can see every change in real-time - great for debugging and demos!

---

## Alternative: NuGet .NET Tool Installation (Secondary)

**For users who prefer package managers or already have .NET installed**

NuGet is a secondary distribution channel. It requires the **.NET 10 Runtime or SDK** to be installed.

```powershell
# Requires .NET 10 Runtime or SDK
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

After installation, configure your MCP client with `"command": "mcp-excel"` (same as standalone exe).

**Update via NuGet:**
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
```

**Uninstall:**
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
```

> **Why NuGet is secondary:** The recommended npm package and the standalone exe require no .NET runtime. NuGet remains available for users who already have .NET installed in their workflow.

---

## Updating the MCP Server

### Check Current Version

```powershell
npx -y @sbroenne/mcp-server-excel@latest --version
```

### Update to New Version

Before restarting or updating, finish the current work and explicitly save and
close the intended workbook sessions. Do not interrupt a refresh or calculation.

| Installation | How to update |
|---|---|
| npm / Copilot plugin | Restart the MCP server/client; `@latest` resolves the server using normal npm caching |
| Claude Desktop MCPB | Restart the server to resolve the npm server; install a new `.mcpb` manually when its configuration changes |
| VS Code extension | Update the extension through VS Code, then restart its bundled server |
| Standalone ZIP | Replace the extracted executable with the new release, then restart the client |
| NuGet | Run `dotnet tool update --global Sbroenne.ExcelMcp.McpServer`, then restart the client |

The npm version command above checks the npm-launched executable, not an
existing server launched by a different installation. A restart does not upgrade
an older binary MCPB or a standalone executable.

**Standalone exe:**

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download the new ZIP: `ExcelMcp-MCP-Server-{version}-windows.zip`
3. Extract and overwrite the existing files in your installation directory

```powershell
# Example update
Expand-Archive "ExcelMcp-MCP-Server-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp" -Force
```

4. Restart your MCP client (VS Code, Claude Desktop, Cursor, etc.)

**NuGet (secondary):**

```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
```

### Check What's New

Before updating, check the [changelog](../CHANGELOG.md) or [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases).

---

## Troubleshooting

### 1. "mcp-excel is not recognized as an internal or external command"

This error applies to the standalone executable. Either use the recommended npm
configuration or add `mcp-excel.exe` to your PATH.

Either:
- Add the directory containing `mcp-excel.exe` to your PATH (see Step 2 above)
- Or use the full path in your MCP client config: `"command": "C:\\Tools\\ExcelMcp\\mcp-excel.exe"`

### 2. MCP Server Not Responding

**Check if the exe exists:**
```powershell
where.exe mcp-excel
# Or with full path:
Test-Path "C:\Tools\ExcelMcp\mcp-excel.exe"
```

**Verify it runs:**
```powershell
npx -y @sbroenne/mcp-server-excel@latest --version
```

### 3. "Workbook is locked" or "Cannot open file"

**Solution:** Close all Excel windows before running ExcelMcp

ExcelMcp requires exclusive access to workbooks (Excel COM limitation).

### 4. MCP Server Still Running Old Version

**Solution:** Check the update steps for your installation method first, then
fully restart your MCP client. `@latest` is resolved at launch; a running process
does not change versions.
- Close VS Code completely (including terminal windows)
- Close Claude Desktop completely
- Reopen the application

---

### 5. Session ID Missing Through a Client Bridge

MCP tool requests use **`session_id`** as the canonical name. Put it directly
in the `arguments` object of `tools/call`, alongside `action`:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "workbook",
    "arguments": {
      "action": "get-info",
      "session_id": "<ID returned by this server>"
    }
  }
}
```

`file open/create`, `file list` entries, and session error context all use
`session_id`. Pass the selected entry's value directly as `session_id`. Never
guess an ID or pick another workbook just because only one is listed.
The legacy `sessionId` input is rejected, even if `session_id` is also present.
CLI JSON continues to use `sessionId`; CLI and MCP sessions are separate.

Use this workflow: list and match the intended workbook; reuse its session or
open/create; operate; list and check that session's `canClose`; close only when
authorized with an explicit `save: true` or `save: false`. No-save close discards
all unsaved edits, including earlier work, and has no tool-level undo.

Calls within one session execute serially, but concurrently submitted requests
have no guaranteed dependency order, and responses can arrive out of order.
Wait for each dependent call before starting the next. Different sessions can
run independently. A `canClose: true` result is a snapshot; do not submit new
work while closing.

If direct local calls work but Cowork or a remote-devices bridge fails, compare
the request received by the server with the request before the bridge. A client
display saying the ID was supplied does not establish what reached the server.
Check the key name, its location, and whether its value is a non-empty string.
Missing-session diagnostics cannot restore an
argument dropped completely by a client.
See [#850](https://github.com/sbroenne/mcp-server-excel/issues/850) and
[#854](https://github.com/sbroenne/mcp-server-excel/issues/854).

Use only a disposable workbook for diagnosis. Do not publish raw logs or real
session IDs, workbook paths, cell contents, or credentials. Share a sanitized
request shape with values replaced, the versions, and whether direct calls work.
MCP and CLI sessions are separate: the CLI cannot close an MCP-owned session.
Avoid opening more sessions while the bridge cannot forward follow-up calls.

---

## Uninstallation

Remove the server entry from your client's configuration as well.

| Installation | How to remove |
|---|---|
| Claude Desktop MCPB | Remove Excel from Claude's Settings > Extensions |
| VS Code extension | Uninstall ExcelMcp through VS Code Extensions |
| Copilot plugin | `copilot plugin uninstall excel-mcp@mcp-server-excel-plugins` |
| npm through npx | Remove the client configuration; no global installation to uninstall |
| Standalone ZIP / NuGet | Use the commands below |

```powershell
# npm: no global installation to remove

# Standalone exe: simply delete the extracted files
Remove-Item "C:\Tools\ExcelMcp\mcp-excel.exe" -Force

# Remove from PATH if you added it
# Settings → System → About → Advanced system settings → Environment Variables
# Edit PATH and remove the ExcelMcp directory

# NuGet (if installed via dotnet tool):
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
```

---

## Getting Help

- **Troubleshooting:** [Troubleshooting](https://excelmcpserver.dev/troubleshooting/) · [FAQ](https://excelmcpserver.dev/faq/)
- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

---

## Next Steps

After installation:

1. **Learn the basics:** Try simple commands like creating worksheets, setting values
2. **Explore features:** See the [Feature Reference](../FEATURES.md) for the complete tool list
3. **Read the guides:**
   - [CLI Installation Guide](INSTALLATION-CLI.md) - for scripting, RPA, and CI/CD
   - [Agent Skills](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-mcp/skills/excel-mcp-report-formatting) - cross-platform AI guidance
4. **Join the community:** Star the repo, report issues, contribute improvements

**Happy automating! 🚀**
