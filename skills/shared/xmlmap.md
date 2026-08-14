# XML Map Reference

Use `xmlmap` for Excel XML maps and in-memory XML import/export.

## Actions

| Action | Purpose | Key parameters |
|--------|---------|----------------|
| `list` | List workbook XML maps | none |
| `add` | Add an XSD schema map | `schema` or `schema_file`; optional `root_element_name`, `map_name` |
| `map-range` | Bind a cell or single-column range to an XPath | `map_name`, `sheet_name`, `range_address`, `xpath`; optional `selection_namespace`, `repeating` |
| `import-xml` | Import XML into an existing map or create an automatically mapped XML table | `xml_data` or `xml_data_file`; either `map_name`, or `sheet_name` plus optional `start_cell` |
| `export-xml` | Return mapped cell values as XML | `map_name` |
| `delete` | Remove a map while leaving existing cell data | `map_name` |

## Import Modes

Use an existing map when XPath mappings already exist:

```text
xmlmap(import-xml, map_name='CustomerMap', xml_data='<customer>...</customer>')
```

Omit `map_name` to let Excel infer a schema, create a map, and create an XML
table at a destination:

```text
xmlmap(import-xml, sheet_name='Sheet1', start_cell='B2', xml_data='<customers>...</customers>')
```

## Security and Determinism

- XML DTDs are rejected.
- XSD `import`, `include`, and `redefine` dependencies are rejected.
- XML `xsi:schemaLocation` and `xsi:noNamespaceSchemaLocation` attributes are
  rejected before Excel can resolve HTTP, UNC, or local-file schemas.
- Use `schema_file` and `xml_data_file` for local file content; the generated
  CLI and MCP surfaces read the file and send its content to Core.
- Import/export stays in memory. URL/file variants that could fetch remote data
  or overwrite server files are intentionally not exposed.
- No dialogs or file pickers are opened.
