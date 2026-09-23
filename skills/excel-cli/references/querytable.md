> **CLI syntax note:** This shared domain guide may use MCP-style `tool(action: ...)` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live `--help`; notably, MCP `file` open/close maps to CLI `session` open/close, and MCP `worksheet` maps to CLI `sheet`.

# querytable - Local Text and Web Imports

Use `querytable` for worksheet QueryTables backed by the desktop Excel COM object model.

## Choose the Right Import Surface

| Need | Tool |
|------|------|
| Direct text/CSV import with delimiter and encoding control | `querytable create-text` |
| Legacy HTML page/table import through Excel's web-query engine | `querytable create-web` |
| Modern connectors, transformations, APIs, JSON, or reusable M | `powerquery` |
| Existing OLEDB/ODBC workbook connection | `connection` |

## Text Import

```text
querytable(action: 'create-text',
    query_table_name: 'OrdersCsv',
    source_path: 'C:\Data\orders.csv',
    sheet_name: 'Orders',
    destination_address: 'A1',
    delimiter: ',',
    text_qualifier: 'double-quote',
    encoding: 65001,
    has_headers: true)
```

- `delimiter` is exactly one character.
- `text_qualifier` is `double-quote`, `single-quote`, or `none`.
- `encoding` is a Windows code page; use `65001` for UTF-8.
- Creation refreshes synchronously so imported data is ready when the call returns.

## Legacy Web Import

```text
querytable(action: 'create-web',
    query_table_name: 'RatesHtml',
    url: 'https://example.com/rates.html',
    sheet_name: 'Rates',
    destination_address: 'A1',
    selection_type: 'specified-tables',
    web_tables: '1',
    formatting: 'none')
```

- `selection_type` is `entire-page`, `all-tables`, or `specified-tables`.
- `web_tables` is required with `specified-tables`.
- `formatting` is `none`, `rich-text`, or `all`.
- This is Excel's legacy HTML web-query engine, not a general HTTP or browser automation API.

## Lifecycle and Refresh

Use `list`, `view`, `set-properties`, `refresh`, `get-refresh-status`, `cancel-refresh`, and `delete` for existing QueryTables.

## Hard Exclusions

Local QueryTable COM automation cannot access Microsoft 365 cloud service state or APIs:

- No workbook sharing or permissions
- No coauthor presence, cursors, conflicts, or live collaboration state
- No comment @mentions, assignments, reactions, or notification delivery
- No authenticated Graph, SharePoint, Teams, or OneDrive service operations
- No Power Query M definition or modern connector configuration

Use the relevant Microsoft 365 service API for cloud workflows and `powerquery` for modern data transformation.
