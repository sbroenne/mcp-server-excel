# Privacy Policy

**Last Updated:** August 28, 2026

## Overview

MCP Server for Excel ("ExcelMcp") is an open-source tool that enables AI assistants to interact with Microsoft Excel. This privacy policy explains how the software handles your data.

## Data Collection Summary

**Telemetry applies to the MCP Server only.** The CLI (`excelcli`) and its background daemon send no telemetry of any kind — the code paths described below exist solely in `ExcelMcp.McpServer`.

ExcelMcp's MCP Server collects **limited, anonymous telemetry** to improve the
software. The statements below describe telemetry collection, not the workbook
data that a tool returns to your chosen AI assistant.

### What We DO Collect (Anonymous Telemetry)

- **Tool usage statistics** - Which tools and actions are used (e.g., "range/get-values")
- **Performance metrics** - How long operations take (duration in milliseconds)
- **Invocation outcome** - Whether an operation succeeded, returned an expected
  negative diagnostic result, or failed
- **Failure class** - A fixed privacy-safe label for input or state, an external
  dependency, timeout or cancellation, Excel runtime, an internal product fault,
  or an unclassified failure
- **Session information** - A random session ID generated each time the server starts
- **Anonymous user ID** - A hashed identifier based on machine identity (not personally identifiable)
- **Application version** - Which version of ExcelMcp is running
- **Unhandled exceptions** - Error type, approved source, and project-owned failure site only
  (never exception messages or stack traces)

### What We DO NOT Collect

- ❌ **File contents** - Workbook content is never included in telemetry
- ❌ **File names or paths** - File paths are not included in telemetry
- ❌ **Personal information** - No names, emails, or account information
- ❌ **Spreadsheet data** - Cell values, formulas, and query results are never included in telemetry
- ❌ **Prompts or messages** - Requests and conversations are never transmitted
- ❌ **Host or client details** - Working directories and MCP client names are not collected
- ❌ **User accounts** - No registration or sign-in required
- ❌ **Error details** - Error messages, response content, exception names, and
  stack traces are not included in invocation outcome telemetry

### Purpose of Telemetry

We use anonymous telemetry to:
- Understand which features are most used
- Identify and fix performance issues
- Prioritize development of new features
- Detect and fix bugs

Invocation telemetry reads only the structured category produced by ExcelMcp; it
does not inspect response or error text. Unknown or missing categories remain
**unclassified**.

### Telemetry Infrastructure

Telemetry is sent to **Azure Application Insights**, a Microsoft service. Data is:
- Transmitted over HTTPS
- Stored in accordance with Microsoft's data handling policies
- Retained for analytics purposes only
- Filtered at ingestion so exception records without ExcelMcp's explicit
  sanitization marker are discarded
- Filtered at ingestion so framework trace logs are discarded in full

## How It Works

ExcelMcp drives the Excel application on your local machine:

1. **Local Processing** - All Excel operations are performed locally via Microsoft's COM API
2. **Your Files Stay Local** - Excel files are read from and written to your local filesystem only
3. **Tool Results** - Data requested by your AI assistant is returned through
   your MCP client or CLI process; your assistant's privacy policy governs how
   it handles that data
4. **Optional Network Features** - Remote M/DAX formatting and Python in Excel
   use external services only when you request those features
5. **Telemetry** - Release builds of the MCP Server can send the anonymous usage
   metrics listed above to Azure Application Insights; the CLI sends no
   telemetry

## Data Flow

When you use ExcelMcp with an AI assistant (like Claude):

1. You send a request to the AI assistant
2. The AI assistant calls ExcelMcp tools on your local machine
3. ExcelMcp performs the requested Excel operations locally
4. Requested results are returned to the AI assistant through your MCP client
5. The MCP Server can send anonymous usage telemetry to Azure Application Insights

Some operations have additional data flows:

- Setting `formatMCode=true` sends the supplied M code to
  [powerqueryformatter.com](https://powerqueryformatter.com/).
- Setting `formatDax=true` sends the supplied DAX formula to
  [DAX Formatter](https://www.daxformatter.com/).
- Python in Excel sends Python code and referenced worksheet data to Microsoft's
  cloud execution environment. This feature requires Microsoft 365, internet
  access, and Python in Excel to be enabled.

Remote formatting is disabled by default and requires explicit consent.
ExcelMcp preserves M and DAX code locally unless you opt in.

**Note:** The AI assistant you use (for example, Claude or GitHub Copilot) has
its own privacy policy governing conversations, tool calls, and returned
workbook data.

## Third-Party Services

- **Azure Application Insights** - Anonymous telemetry is sent to this Microsoft service. See [Microsoft's Privacy Statement](https://privacy.microsoft.com/privacystatement).
- **Microsoft Excel** - ExcelMcp requires Microsoft Excel installed on your machine. Excel is subject to Microsoft's privacy policy.
- **Microsoft Python in Excel** - Python code and referenced worksheet data run
  in Microsoft's cloud when you use the Python in Excel operations.
- **Power Query Formatter** - Receives M code only when you set
  `formatMCode=true`.
- **DAX Formatter** - Receives DAX formulas only when you set
  `formatDax=true`.
- **AI Assistants** - When used with AI assistants like Claude, those services have their own privacy policies.

## Open Source

ExcelMcp is open source software. You can review the complete source code at:
https://github.com/sbroenne/mcp-server-excel

## Security

- ExcelMcp runs with the same permissions as your user account
- It can only access files and Excel instances that your user account can access
- No elevated privileges are required or requested

## Children's Privacy

ExcelMcp does not knowingly collect any information from anyone, including children under 13 years of age.

## Changes to This Policy

If we make changes to this privacy policy, we will update the "Last Updated" date above and publish the updated policy in our GitHub repository.

## Contact

For questions about this privacy policy or the ExcelMcp project:

- **GitHub Issues:** https://github.com/sbroenne/mcp-server-excel/issues
- **Repository:** https://github.com/sbroenne/mcp-server-excel

---

**Summary:** ExcelMcp drives local Excel and reads or writes workbook files on
your machine. Requested tool results can be returned to your chosen AI
assistant. Remote code formatting and Python in Excel use external services only
when requested. The MCP Server can send anonymous usage telemetry, but that
telemetry excludes workbook contents, file names, paths, and personal
information. The CLI sends no telemetry.
