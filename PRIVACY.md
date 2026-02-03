# Privacy Policy

**Last Updated:** January 13, 2026

## Overview

MCP Server for Excel ("ExcelMcp") is an open-source tool that enables AI assistants to interact with Microsoft Excel. This privacy policy explains how the software handles your data.

## Data Collection Summary

ExcelMcp collects **limited, anonymous telemetry** to improve the software. Here's what we do and don't collect:

### What We DO Collect (Anonymous Telemetry)

- **Tool usage statistics** - Which tools and actions are used (e.g., "excel_range/get-values")
- **Performance metrics** - How long operations take (duration in milliseconds)
- **Success/failure rates** - Whether operations completed successfully
- **Session information** - A random session ID generated each time the server starts
- **Anonymous user ID** - A hashed identifier based on machine identity (not personally identifiable)
- **Application version** - Which version of ExcelMcp is running
- **Unhandled exceptions** - Error types (not error messages or stack traces with sensitive data)

### What We DO NOT Collect

- **File contents** - We never collect data from your Excel files
- **File names or paths** - File paths are hashed locally; actual paths are never transmitted
- **Personal information** - No names, emails, or account information
- **Spreadsheet data** - Cell values, formulas, and data remain completely private
- **User accounts** - No registration or sign-in required

### Purpose of Telemetry

We use anonymous telemetry to:
- Understand which features are most used
- Identify and fix performance issues
- Prioritize development of new features
- Detect and fix bugs

### Telemetry Infrastructure

Telemetry is sent to **Azure Application Insights**, a Microsoft service. Data is:
- Transmitted over HTTPS
- Stored in accordance with Microsoft's data handling policies
- Retained for analytics purposes only

## How It Works

ExcelMcp operates on your local machine:

1. **Local Processing** - All Excel operations are performed locally via Microsoft's COM API
2. **Your Files Stay Local** - Excel files are read from and written to your local filesystem only
3. **Minimal Network Usage** - The only network traffic is anonymous telemetry to Azure Application Insights

## Data Flow

When you use ExcelMcp with an AI assistant (like Claude):

1. You send a request to the AI assistant
2. The AI assistant calls ExcelMcp tools on your local machine
3. ExcelMcp performs the requested Excel operations locally
4. Anonymous usage telemetry is sent to Azure Application Insights
5. Results are returned to the AI assistant

**Note:** The AI assistant you use (e.g., Claude) has its own privacy policy governing how it handles your conversations and data. ExcelMcp only handles the local Excel operations and sends anonymous usage metrics.

## Third-Party Services

- **Azure Application Insights** - Anonymous telemetry is sent to this Microsoft service. See [Microsoft's Privacy Statement](https://privacy.microsoft.com/privacystatement).
- **Microsoft Excel** - ExcelMcp requires Microsoft Excel installed on your machine. Excel is subject to Microsoft's privacy policy.
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

**Summary:** ExcelMcp processes your Excel files locally on your machine. We collect anonymous usage telemetry (tool usage, performance, errors) to improve the software, but never collect your file contents, file names, or personal information.
