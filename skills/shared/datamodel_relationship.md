# datamodel_relationship - Table Relationship Operations

**Actions**: list-relationships, create-relationship, read-relationship, update-relationship, delete-relationship

**When to use datamodel_relationship**:
- Creating a star schema (linking fact tables to dimension/lookup tables)
- Enabling cross-table DAX calculations (e.g., RELATED(), RELATEDTABLE())
- Building Data Models that use data from multiple tables
- Use `datamodel` for tables and DAX measures, NOT for relationships

**PREREQUISITE: Both tables must be in the Data Model first**

Use `table(action:'add-to-data-model')` or `powerquery(loadDestination:'data-model')` before creating relationships.

**Create relationship pattern** (star schema example):

```
# Orders[ProductID] (many) → Products[ProductID] (one)
datamodel_relationship(action:'create-relationship',
    fromTable:'Orders', fromColumn:'ProductID',   # many-side (foreign key, detail table)
    toTable:'Products', toColumn:'ProductID')      # one-side (primary key, lookup table)
```

**fromTable/fromColumn = many-side (detail table, foreign key)**
**toTable/toColumn = one-side (lookup table, primary key)**

**Action guide**:
- `create-relationship`: Link two tables on a common column. Specify from (many) → to (one).
- `list-relationships`: View all existing relationships before schema changes.
- `read-relationship`: Get details of a specific relationship by its table/column identifiers.
- `update-relationship`: Toggle active/inactive state (active=true/false).
- `delete-relationship`: Remove a relationship. WARNING: deleting/recreating tables also deletes their relationships.

**ACTIVE vs INACTIVE relationships**:
- Only ONE active relationship can exist between two table pairs
- Use `active=false` for alternative join paths
- Activate alternative paths in DAX with `USERELATIONSHIP()`

**Common mistakes**:
- Creating relationship before adding tables to Data Model (fails with "table not found")
- Swapping fromTable/toTable (many/one sides matter for RELATED() to work correctly)
- Forgetting that DELETE TABLE removes all relationships on that table
