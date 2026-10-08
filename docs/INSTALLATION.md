# Installation Guide - ExcelMcp

Install ExcelMcp to automate installed Microsoft Excel on Windows from an AI
assistant or the command line. It ships two **equal entry points** — the
**MCP Server** for MCP clients and the **CLI** (`excelcli`) for coding agents
and scripts. Choose the guide for your workflow; the entry points are
independent:

| Guide | Best For |
|-------|----------|
| 📖 **[Installing the MCP Server](INSTALLATION-MCP-SERVER.md)** | AI assistants — GitHub Copilot, Claude Desktop, Cursor, Windsurf, and any other MCP client |
| 📖 **[Installing the CLI](INSTALLATION-CLI.md)** | Scripting, RPA, and coding agents on a desktop Excel host |

Both entry points support **Windows with Microsoft Excel 2016+** and
**Apple Silicon macOS with Excel for Mac 16.112+**. Windows provides the complete
operation set; macOS support is **experimental beta** with a
[capability-gated subset and explicit limitations](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta).
The standalone distributions do not require a .NET runtime.

| Where you work | Recommended installation |
|---|---|
| VS Code with GitHub Copilot | VS Code extension; bundles the server and its skill |
| Claude Desktop | [Claude Desktop setup guide](guides/CLAUDE-DESKTOP.md); MCPB configures direct npx with `@latest` (Node.js/npm required) |
| Another MCP client | npm through `npx -y @sbroenne/mcp-server-excel@latest` |
| Coding agents and scripts | `npx -y @sbroenne/excelcli@latest`, or global npm for a command on PATH |
| No npm downloads desired | Standalone ZIP; replace the executable manually for updates |

Windows and Apple Silicon macOS use separate native archives; one executable
file cannot be shared across PE/Windows and Mach-O/macOS. Intel macOS is
unsupported and fails closed rather than selecting the ARM64 runtime. See
[macOS distribution readiness](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/MACOS-DISTRIBUTION.md)
for package inspection, signing, and notarization details.

The macOS backend does not require an Office.js add-in, hosted manifest, or
localhost certificate. Only actions enabled in the generated macOS inventory
are supported.

### Optional macOS native helper

Helper artifact acceptance and the first independent release are still pending.
Native actions work without it; installing a helper does not enable gated
Power Query, formula, or other unverified actions.

When a separate `helper-vX.Y.Z` release provides the artifact:

1. Download `ExcelMcpHelper.xlam` and `SHA256SUMS` from that release in
   [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases).
   Verify the checksum. Keep the filename `ExcelMcpHelper.xlam`.
2. Put it in a stable location you control. In Excel, open **Tools > Excel Add-ins**,
   browse to the downloaded file, and enable it.
3. Approve macros interactively only if you trust that release. ExcelMcp does
   not alter macro security, click trust dialogs, or install source modules.
4. With desktop Excel open, check the installed helper through the CLI:

   ```powershell
   '{"command":"service.helper-check"}' | excelcli -q batch
   ```

The check returns the helper's `version` and `primitives`, or an explicit missing,
incompatible, malformed, permission, or execution error. The server accepts
helper major **1** and checks each requested primitive before helper-backed
mutation. The helper and server release numbers do **not** need to match.
Compatible older helpers remain usable for primitives they supply.

Users install the released `.xlam`; they do not import `.bas` files or build it.
Maintainer bootstrap/build instructions are in
[the helper source README](https://github.com/sbroenne/mcp-server-excel/blob/main/helper/mac/README.md).

> **Tip:** The **VS Code Extension** bundles the MCP Server only (install the CLI separately if you need it for scripting). The **GitHub Copilot plugins** are separate — install `excel-mcp` and/or `excel-cli` depending on which entry point you need — see the MCP Server guide's Quick Start for the one-click paths.

The Copilot plugins are also listed in
[Awesome Copilot](https://github.com/github/awesome-copilot), the default
marketplace in current Copilot clients. Each installation guide includes that
install route and our direct marketplace alternative.

---

## Agent Skills Installation (Cross-Platform)

**Best for:** Optional presentation guidance for requested Excel reports.

The VS Code extension registers `excel-mcp-report-formatting`. Plugins remain
`excel-mcp` and `excel-cli`; their contained skills have narrower names and scope.
After the release containing this change is published, install directly with:

```powershell
# CLI report-formatting skill
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting

# MCP report-formatting skill
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp-report-formatting

# Interactive install - select one or both formatting skills
npx skills add sbroenne/mcp-server-excel-plugins

# Install for specific agents
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting -a cursor
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp-report-formatting -a claude-code

# Install both skills
npx skills add sbroenne/mcp-server-excel-plugins --skill '*'

# Install globally (user-wide)
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting --global
```

**Supports 43+ agents** including claude-code, github-copilot, cursor, windsurf, gemini-cli, codex, goose, cline, continue, replit, and more.

**Manual Installation:**

Existing broad `excel-cli` and `excel-mcp` standalone skills remain installed
until removed with the client's skill manager. Keep plugin/server configuration.
Use complete prepared packages, not source entries without their references.

1. Download `excel-skills-v{version}.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. The package contains both skills:
   - `skills/excel-cli-report-formatting/` - presentation through `excelcli`
   - `skills/excel-mcp-report-formatting/` - presentation through MCP tools
3. Extract the skill(s) you need to your AI assistant's skills directory:
   - Copilot: `~/.copilot/skills/<skill-name>/`
   - Claude Code: `.claude/skills/<skill-name>/`
   - Cursor: `.cursor/skills/<skill-name>/`

**See:** [Agent Skills Documentation](../docs/AGENT-SKILLS.md)

---

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

**Happy automating! 🚀**
