# Local text and web imports

Choose the import workflow by the source and required transformation, not
simply because every option can place data on a worksheet.

Current commands and inputs come from CLI help or MCP tool descriptions.

## Choose the Right Import Surface

| Need | Approach |
|------|----------|
| Direct text/CSV import with control over interpretation | Worksheet QueryTable |
| Legacy HTML page or table import | Excel's legacy web-query engine |
| Modern connectors, transformations, APIs, or reusable M | [Power Query](powerquery.md) |
| Existing OLEDB/ODBC source | Workbook connection with its installed provider/driver |

## Text Import

Use the intended readable source file and inspect the destination first.
Choose its delimiter, quoting, encoding, and header interpretation deliberately;
incorrect choices can alter dates, identifiers, and column boundaries.

Creation loads synchronously, so read the resulting cells before relying on
the imported data. A completed call does not establish that Excel interpreted
every field as the business intended.

## Legacy Web Import

Use the user's intended HTML page, not a guessed source. Choose whether the
task needs the whole page or selected tables.

This is Excel's legacy web-query engine, not a general browser or a modern
authenticated cloud connector. Use Power Query for modern transformations
and source-specific connectors.

## Lifecycle and Refresh

Inspect the existing import's destination and refresh behavior before changing
it. Refresh executes the source again and can replace loaded data; it is not
merely inspecting a frozen snapshot.

If refresh fails or is cancelled, inspect the resulting state before retrying.
Do not assume the destination remained unchanged.

## Hard Exclusions

Local QueryTable automation does not expose Microsoft 365 sharing permissions,
coauthor presence, comment mentions/reactions, or service notification delivery.
Use the relevant Microsoft 365 API for those workflows, not worksheet import
as a service-access workaround.
