# DMV Query Reference (Excel's Embedded Analysis Services)

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

## Full TMSCHEMA Catalog

| Category | DMVs |
|----------|------|
| Structure | TMSCHEMA_MODEL, TMSCHEMA_TABLES, TMSCHEMA_COLUMNS, TMSCHEMA_HIERARCHIES, TMSCHEMA_LEVELS |
| Measures | TMSCHEMA_MEASURES, TMSCHEMA_KPIS, TMSCHEMA_FORMAT_STRING_DEFINITIONS |
| Relationships | TMSCHEMA_RELATIONSHIPS |
| Security | TMSCHEMA_ROLES, TMSCHEMA_ROLE_MEMBERSHIPS, TMSCHEMA_TABLE_PERMISSIONS, TMSCHEMA_COLUMN_PERMISSIONS |
| Partitions | TMSCHEMA_PARTITIONS, TMSCHEMA_DATA_SOURCES |
| Metadata | TMSCHEMA_ANNOTATIONS, TMSCHEMA_EXTENDED_PROPERTIES, TMSCHEMA_CULTURES, TMSCHEMA_OBJECT_TRANSLATIONS |
| Perspectives | TMSCHEMA_PERSPECTIVES, TMSCHEMA_PERSPECTIVE_TABLES, TMSCHEMA_PERSPECTIVE_COLUMNS, TMSCHEMA_PERSPECTIVE_MEASURES |
| Calculations | TMSCHEMA_CALCULATION_GROUPS, TMSCHEMA_CALCULATION_ITEMS, TMSCHEMA_EXPRESSIONS |

Reference: [Microsoft DMV Docs](https://learn.microsoft.com/en-us/analysis-services/instances/use-dynamic-management-views-dmvs-to-monitor-analysis-services)
