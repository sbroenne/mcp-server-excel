# XML Map Reference

Use `xmlmap` for Excel XML maps and in-memory XML import/export.

Use tool schemas/native help for action inputs. Deleting a map leaves existing
cell data; it does not clear the imported worksheet.

## Import Modes

Use an existing map when XPath mappings already exist:

```mcp
xmlmap(action: 'import-xml', session_id: sessionId, map_name: 'CustomerMap', xml_data_file: 'customers.xml')
```

```cli
excelcli -q xmlmap import-xml --session $sessionId --map-name CustomerMap --xml-data-file customers.xml
```

Omit `map_name` (MCP) / `--map-name` (CLI) to let Excel infer a schema, create a map, and create an XML
table at a destination:

```mcp
xmlmap(action: 'import-xml', session_id: sessionId, sheet_name: 'Sheet1', start_cell: 'B2', xml_data_file: 'customers.xml')
```

```cli
excelcli -q xmlmap import-xml --session $sessionId --sheet Sheet1 --start-cell B2 --xml-data-file customers.xml
```

These alternatives assume a captured session and a known readable XML file.
Inspect the existing map or empty destination cells before choosing one.

## Security and Determinism

- XML DTDs are rejected.
- XSD `import`, `include`, and `redefine` dependencies are rejected.
- XML `xsi:schemaLocation` and `xsi:noNamespaceSchemaLocation` attributes are
  rejected before Excel can resolve HTTP, UNC, or local-file schemas.
- Use `schema_file` and `xml_data_file` (MCP) / `--schema-file` and
  `--xml-data-file` (CLI) for local file content; the generated
  CLI and MCP surfaces read the file and send its content to Core.
- Import/export stays in memory. URL/file variants that could fetch remote data
  or overwrite server files are intentionally not exposed.
- No dialogs or file pickers are opened.
