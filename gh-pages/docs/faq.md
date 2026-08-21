---
title: Frequently Asked Questions
description: >-
  Answers to common questions about Excel MCP Server - Windows and Excel
  requirements, CLI vs. MCP Server, workbook safety, cost, and privacy.
keywords: "Excel MCP FAQ, does Excel MCP need Excel installed, Excel MCP Windows only, Excel MCP vs openpyxl, Excel MCP CLI or MCP Server, is Excel MCP free"
---

# Frequently Asked Questions

<!--
  Every `###` heading on this page is parsed by gh-pages/hooks.py into FAQPage
  JSON-LD (see `_faq_jsonld`). The structured data is derived from this visible
  content, so there is nothing to keep in sync by hand - just write the
  questions here. Headings rather than collapsible admonitions are used
  deliberately: they give each answer a stable anchor that can be deep-linked
  from another page or straight from a search result.
-->

Hitting an actual error rather than a question? See
[Troubleshooting](troubleshooting.md).

## Getting started

### Do I need to know how Excel automation works to use this?

No. You talk to your AI assistant in plain language ("build a PivotTable of
sales by product and chart it") and it drives Excel for you. The
[feature reference](features.md) is there when you want to see everything that's
possible - you don't need to memorize it.

### CLI or MCP Server - which should I install?

Both expose the **same 325 operations**. Use the **MCP Server** for
conversational AI (Claude Desktop, VS Code Chat); use the **CLI** (`excelcli`)
for coding agents and scripting, where it uses ~64% fewer tokens. You can
install both. See [Installation](installation.md).

### Which AI assistants does it work with?

Any client that speaks the Model Context Protocol - GitHub Copilot in VS Code,
Claude Desktop, Claude Code, Cursor, and others. Coding agents that can run
shell commands can use the CLI directly instead. See
[MCP Server installation](installation-mcp-server.md).

## Requirements and platform

### Does it require Microsoft Excel to be installed?

Yes. Excel MCP Server drives the **real Excel application** through its COM API,
so it's **Windows-only** and needs **Excel 2016 or later** installed locally. It
is not a file-format parser and does not run on macOS or Linux.

### Why drive real Excel instead of parsing the file?

Because Excel itself is the only thing that fully understands its own file
format. Recalculation, Power Query evaluation, the Data Model, PivotTable
caches, and chart rendering all live in the application, not in the `.xlsx`
container. See
[Excel automation vs. file parsers](guides/excel-automation-vs-file-parsers.md)
for the full comparison with tools like openpyxl and pandas.

### Will it damage my existing workbooks?

No. Excel itself opens and saves the file, so formulas, PivotTables, charts,
macros, the Data Model, and formatting are all preserved. Other tools that
rewrite the `.xlsx` file directly can silently drop those; here Excel does the
work.

## Cost, privacy and support

### Does it cost anything or send my data anywhere?

Excel MCP Server is free and open source (MIT). It runs locally against your own
Excel. A few opt-in features reach the internet (remote M/DAX formatting, and
Python in Excel, which runs in Microsoft's cloud). See
[Privacy](privacy.md) for details.

### Where do I report a bug or ask for a feature?

On [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues). If
something is failing rather than missing, check
[Troubleshooting](troubleshooting.md) first - most first-time failures are one
of a handful of known setup issues.
