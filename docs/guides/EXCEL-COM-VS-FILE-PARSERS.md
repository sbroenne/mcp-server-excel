# Real Excel Automation vs. File-Parser Libraries

There are two fundamentally different ways to automate Excel files, and the choice
determines what is possible. Neither is universally better — they solve different
problems.

## The two approaches

**File parsers** read and write the `.xlsx` package directly. An `.xlsx` file is a
ZIP archive of XML, so a library can open it, edit the XML, and write it back
without Excel being installed. `openpyxl`, `ExcelJS`, `SheetJS`, `EPPlus`, and
`ClosedXML` all work this way.

**COM automation** launches the real Microsoft Excel application and drives it
through Excel's official `Excel.Application` API — the same API VBA uses. ExcelMcp
takes this approach.

## What each can do

| Capability | File parsers | COM automation |
|---|---|---|
| Read and write cell values | Yes | Yes |
| Read and write formulas (as text) | Yes | Yes |
| **Calculate** formula results | No — values are stale until Excel opens the file | Yes, Excel's own engine |
| Cell formatting, styles, number formats | Yes | Yes |
| Create charts | Basic | Full Excel chart engine |
| **Refresh** Power Query | No | Yes |
| **Refresh** PivotTables and the Data Model | No | Yes |
| Evaluate DAX / query the Data Model | No | Yes |
| Run VBA macros | No | Yes |
| Run Python `=PY()` formulas | No | Yes |
| Preserve unknown/complex workbook parts | Varies — some are dropped on rewrite | Yes, Excel owns the file |
| Interactive authentication for protected sources | No | Yes |
| Runs on Linux / macOS / containers | Yes | No — Windows + Excel required |
| Runs without Excel installed | Yes | No |
| Speed for bulk cell writes | Very fast | Slower (process boundary) |

## Why the difference exists

A file parser sees the *stored* state of a workbook. Formula results, PivotTable
caches, and query results are all values Excel wrote the last time it calculated.
A parser can change a formula's text, but it cannot produce the new result — only
Excel's calculation engine can. The same applies to Power Query (an engine inside
Excel), the Data Model (an Analysis Services tabular engine embedded in Excel),
and VBA (a runtime hosted by Excel).

Rewriting a workbook package also risks losing parts the library does not model.
When Excel itself saves the file, everything it does not touch is preserved by
construction.

## Which one should you use?

**Use a file parser when:**

- You need to run on Linux, macOS, or in a container
- Excel is not installed and cannot be
- You are generating simple reports from scratch — values, formatting, basic charts
- You are writing large volumes of cell data and throughput matters
- The workbook has no Power Query, PivotTables, Data Model, or macros

**Use COM automation (ExcelMcp) when:**

- The workbook contains Power Query, PivotTables, the Data Model, or VBA
- You need calculated formula results, not just formula text
- You must preserve an existing complex workbook exactly
- The data source needs interactive sign-in
- You want an AI assistant to work with real business workbooks that already exist

The dividing line in practice: **generating a new simple file** favours parsers;
**operating on an existing real-world workbook** favours COM.

## What ExcelMcp adds on top of COM

Raw COM automation from a script is possible but unpleasant — STA threading, COM
object lifetime, message filters, and Excel process cleanup are all easy to get
wrong and leak `EXCEL.EXE` processes. ExcelMcp handles that layer and exposes 325
operations across 31 tools through two equal entry points:

- an **MCP server** for conversational AI clients (Claude, Copilot, Cursor)
- a **CLI** (`excelcli`) for scripting and coding agents

The CLI additionally avoids loading MCP tool schemas into an agent's context. In a
same-task, same-model benchmark the CLI workflow used about 59K tokens versus 163K
for MCP — a 64% reduction. Actual usage varies by client, model, and workflow.

## Requirements and trade-offs

ExcelMcp is **Windows-only and requires Microsoft Excel desktop (2016 or later)**.
That is the direct cost of using Excel's real engines. If you need
cross-platform execution, a file parser is the right tool and no amount of
architecture changes that.

## Related

- [Architecture](../ARCHITECTURE.md)
- [Feature overview](../../FEATURES.md)
- [Installation](../INSTALLATION.md)
