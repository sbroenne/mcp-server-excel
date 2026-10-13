# Privacy Policy

**Last Updated:** October 13, 2026

## Overview

MCP Server for Excel ("ExcelMcp") is an open-source tool that enables AI assistants to interact with Microsoft Excel. This privacy policy explains how the software handles your data.

## Data Collection Summary

ExcelMcp's MCP Server and CLI collect **limited, pseudonymous telemetry** to improve
the software. The CLI background daemon does not send telemetry independently;
the short-lived `excelcli` command process records operations requested through
it. The statements below describe telemetry collection, not the workbook data
that a tool returns to your chosen AI assistant.

### What We DO Collect (Pseudonymous Telemetry)

- **Tool usage statistics** - Which tools and actions are used (e.g., "range/get-values").
  Only known ExcelMcp commands are reported; anything else is reported as "other/other"
- **Entry point** - Whether the operation was requested through the CLI or MCP Server
- **Performance metrics** - How long operations take (duration in milliseconds)
- **Invocation outcome** - Whether an operation succeeded, returned an expected
  negative diagnostic result, or failed
- **Failure class** - A fixed privacy-safe label for input or state, an external
  dependency, timeout or cancellation, Excel runtime, an internal product fault,
  or an unclassified failure
- **Session information** - A random telemetry session ID generated each time the CLI or server starts
- **Pseudonymous user ID** - A random ID stored at
  `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id`. The MCP Server and CLI share it
  for the same Windows user profile, and it persists across restarts until the
  file is deleted. The file is created on first use even if telemetry is off.
  If the file cannot be read or written, that process uses a temporary random
  ID instead.
- **One-time migration link** - If telemetry is active when the ID file is
  first created, one `UserIdMigration` event is sent to the active destination
  with the new ID and the previous 16-character, unsalted hash of the machine
  name, user name, and platform. The raw names are not sent. The hash is
  included only to connect older usage events to the new ID; weekly reports do
  not publish it.
- **Application version** - Which version of ExcelMcp is running
- **Pseudonymous instance label** - `instance-` followed by the first eight
  characters of the random user ID. CLI and MCP events, requests, and duration
  metrics use this label instead of the computer name.
- **Unhandled exceptions (MCP Server)** - Error type names, approved source, and
  project-owned failure site only (never exception messages or stack traces)

### What We DO NOT Collect

- ❌ **File contents** - Workbook content is never included in telemetry
- ❌ **File names or paths** - File paths are not included in telemetry
- ❌ **Personal information** - No raw names, emails, or account information
- ❌ **Spreadsheet data** - Cell values, formulas, and query results are never included in telemetry
- ❌ **Prompts or messages** - Requests and conversations are never transmitted
- ❌ **Host or client details** - Computer names, working directories, and MCP client names are not collected
- ❌ **User accounts** - No registration or sign-in required
- ❌ **Error details** - Error messages, response content, and stack traces are
  never included. Exception type names appear only in the unhandled-exception
  reports listed above, not in invocation outcome telemetry
- ❌ **Workbook session IDs or tool arguments** - Telemetry never includes the
  Excel workbook `workbook_session_id` (MCP) / `sessionId` (CLI) value, workbook path, request
  arguments, or user content. The random process telemetry session ID listed
  above is separate from workbook identity.

### Purpose of Telemetry

We use pseudonymous telemetry to:
- Understand which features are most used
- Identify and fix performance issues
- Prioritize development of new features
- Detect and fix bugs

Invocation telemetry reads only the structured category produced by ExcelMcp; it
does not inspect response or error text. Unknown or missing categories remain
**unclassified**.

### Telemetry Infrastructure

By default, released packages with an embedded connection string send telemetry
to **Azure Application Insights**, a Microsoft service. Data is:
- Transmitted over HTTPS
- Stored in accordance with Microsoft's data handling policies
- Retained for analytics purposes only
- Filtered at ingestion so exception records without ExcelMcp's explicit
  sanitization marker are discarded
- Filtered at ingestion so framework trace logs are discarded in full

Set `EXCELMCP_TELEMETRY_MODE=off` to disable telemetry collection and delivery in
either the CLI or MCP Server. Contributor builds without an embedded connection
string default to off.

Optional `EXCELMCP_TELEMETRY_MODE=local` sends only the approved operation
telemetry above to a loopback OpenTelemetry receiver on your own machine. It
overrides the embedded Azure destination and does not create an Azure telemetry
client. It does not enable framework logs, raw exception details, or additional
workbook data collection. Access to the local dashboard and any data it retains
is your responsibility. See [local telemetry setup](docs/DEVELOPMENT.md#local-development-with-telemetry).

If telemetry is off when the ID is created, no migration link is sent or
deferred. If it is in local mode, the link goes only to the local receiver and
is not deferred for Azure. In either case, older Azure events may remain under
the previous ID in weekly reports. Delete
`%LOCALAPPDATA%\ExcelMcp\telemetry-user-id` and restart the MCP Server to reset
the stored ID; the next active telemetry use creates a new one.

## How It Works

ExcelMcp drives the Excel application on your local machine:

1. **Local Processing** - All Excel operations are performed locally via Microsoft's COM API
2. **Your Files Stay Local** - Excel files are read from and written to your local filesystem only
3. **Tool Results** - Data requested by your AI assistant is returned through
   your MCP client or CLI process; your assistant's privacy policy governs how
   it handles that data
4. **Optional Network Features** - Remote M/DAX formatting and Python in Excel
   use external services only when you request those features
5. **Telemetry** - Release builds of the MCP Server and CLI can send the
   pseudonymous usage metrics listed above to Azure Application Insights

## Data Flow

When you use ExcelMcp with an AI assistant (like Claude):

1. You send a request to the AI assistant
2. The AI assistant calls ExcelMcp tools on your local machine
3. ExcelMcp performs the requested Excel operations locally
4. Requested results are returned to the AI assistant through your MCP client
5. The MCP Server or CLI can send pseudonymous usage telemetry to Azure Application Insights

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

- **Azure Application Insights** - Pseudonymous telemetry is sent to this Microsoft service. See [Microsoft's Privacy Statement](https://privacy.microsoft.com/privacystatement).
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
when requested. The MCP Server and CLI can send pseudonymous usage telemetry.
It excludes workbook contents, file names, paths, arguments, and raw machine or
user names. A one-time migration event may include the previous name-derived
hash so older usage can be counted under the new random ID.
