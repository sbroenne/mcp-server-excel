# DMV Query Reference (Excel's Embedded Analysis Services)

**Windows-only in the experimental macOS beta.** `datamodel execute-dmv` and
the other model actions are unavailable on Mac. See
[macOS support](https://excelmcpserver.dev/macos-support/).

## When to Use DMV Queries

Use DMV queries when ordinary model inspection cannot answer a metadata
question. Use current CLI help or MCP tool descriptions for executing a query;
the guidance here concerns Excel's embedded provider and query language.

| Use Case | DMV to Use |
|----------|-----------|
| Measure metadata beyond regular list/read operations | `TMSCHEMA_MEASURES` |
| Additional relationship metadata | `TMSCHEMA_RELATIONSHIPS` |
| Impact analysis — what depends on a measure/column | `DISCOVER_CALC_DEPENDENCY` |
| List all available DMV views on this workbook | `DISCOVER_SCHEMA_ROWSETS` |

**Do NOT use DMV queries for:**
- Reading regular worksheet data - use worksheet value reads
- Listing Power Query queries - inspect the workbook's stored queries
- Reading PivotTable results - inspect the actual summary data

SYNTAX: `SELECT * FROM $SYSTEM.<SchemaRowset>`

LIMITATIONS:
- ONLY `SELECT *` works — specific column selection fails
- Some TMSCHEMA views return empty results in Excel's embedded AS

## Working DMV Queries (verified)

| Query | Returns |
|-------|---------|
| `SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES` | All DAX measures with formulas |
| `SELECT * FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS` | All relationships between tables |
| `SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY` | Calculation dependencies (impact analysis) |
| `SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS` | Database/catalog metadata |
| `SELECT * FROM $SYSTEM.DISCOVER_SCHEMA_ROWSETS` | List all available DMVs |

## May Return Empty in Excel

`TMSCHEMA_TABLES`, `TMSCHEMA_COLUMNS`, `TMSCHEMA_PARTITIONS`

Discover available rowsets in the actual workbook rather than assuming the full
Analysis Services catalog applies to Excel's embedded provider.

Reference: [Microsoft DMV Docs](https://learn.microsoft.com/en-us/analysis-services/instances/use-dynamic-management-views-dmvs-to-monitor-analysis-services)
