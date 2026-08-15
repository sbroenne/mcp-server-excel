# Guides

Task-focused walkthroughs for the things people most often want an AI assistant to
do in Excel. Each guide shows what to ask for, the equivalent CLI commands, how to
verify the result, and the gotchas that bite people.

For the exhaustive operation reference, see the
[feature documentation](../../FEATURES.md).

## Automating data

- **[Refresh Power Query from an AI assistant](REFRESH-POWER-QUERY.md)** —
  refresh one query or all of them, test M code before committing it, and choose
  where the data lands.
- **[Query the Excel Data Model with DAX](QUERY-DATA-MODEL-WITH-DAX.md)** —
  create measures, run DAX queries, inspect the model with DMVs, and manage
  relationships.

## Building reports

- **[Build and update PivotTables with an AI assistant](AUTOMATE-PIVOTTABLES.md)** —
  choose the right source, configure fields, chart the result, and keep everything
  refreshed.

## Working with existing workbooks

- **[Run VBA macros from an AI agent](RUN-VBA-MACROS.md)** — enable VBA trust,
  inspect modules, execute macros, and import new code.

## Choosing an approach

- **[Real Excel automation vs. file-parser libraries](EXCEL-COM-VS-FILE-PARSERS.md)** —
  what COM automation can do that `openpyxl`, `ExcelJS`, and `EPPlus` cannot, and
  when a file parser is the better choice.
