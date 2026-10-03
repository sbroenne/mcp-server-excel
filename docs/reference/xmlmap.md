# XML maps

XML maps connect structured XML paths to worksheet cells. Use current CLI
help or MCP tool descriptions for commands and supported inputs.

## Import Modes

Reuse an existing map when its schema and cell mappings describe the intended
data. Otherwise Excel can infer a schema and create a map and worksheet Table
at a chosen destination.

Inspect the existing mappings or destination cells before importing. Import
can replace worksheet data; removing a map does not clear previously imported
cells. Export returns mapped data, not arbitrary workbook contents.

## Security and Determinism

Use self-contained schemas and XML from the intended source. DTDs, external
schema dependencies, and schema-location attributes are rejected before Excel
can resolve unexpected web, network, or local-file resources.

Local file inputs are read as content for the operation; they do not authorize
Excel to fetch referenced resources. XML exchange avoids dialogs and remote
URL/file variants that could fetch data or overwrite files unexpectedly.
