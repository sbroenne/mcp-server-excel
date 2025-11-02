# Excel Naming Best Practices

**When creating Excel objects, follow these community-standard naming conventions:**

## Quick Rules

✅ **DO:**
- Use **PascalCase** for tables: `SalesData`, `CustomerList`
- Use **UPPER_SNAKE_CASE** for constants: `TAX_RATE`, `MAX_DISCOUNT`
- Be descriptive: `Customer` not `Table1`
- Use singular nouns for tables: `Customer` not `Customers`

❌ **DON'T:**
- Use spaces in table/query names (use underscores instead)
- Use generic names: `Data`, `Input`, `Table1`
- Use special characters: `@`, `#`, `$`, `-` (except underscore)

## Examples by Object Type

**Tables:**
```
✅ SalesData, Customer, OrderHistory
✅ tbl_Sales (with prefix for large workbooks)
❌ Sales Data, Table1, data
```

**Named Ranges (Parameters):**
```
✅ TAX_RATE, START_DATE, DISCOUNT_RATE
✅ prm_TaxRate (with prefix)
❌ rate, x, temp
```

**Power Queries:**
```
✅ Transform_Customer, Load_Sales, CRM_Extract
❌ Query1, Query2, data
```

**Worksheets:**
```
✅ Sales Data, 01_Input, Dashboard (spaces OK here!)
✅ _Calculations (underscore to hide/group)
❌ Sheet1, data, temp
```

**Table Columns:**
```
✅ FirstName, LastName, OrderDate, AmountUSD
❌ Name, Date, Amount, Col1
```

## Optional Prefixes (for large workbooks)

Use type prefixes to instantly identify object types:
- `tbl_Sales` (tables)
- `rng_Criteria` (named ranges)
- `pq_Transform` (Power Queries)
- `prm_StartDate` (parameters)

## Why This Matters

Good names make workbooks:
- **Self-documenting** - others understand your logic
- **Maintainable** - easy to find and update
- **Formula-friendly** - `=SUM(Sales[Amount])` vs `=SUM(Table1[Column1])`
- **LLM-friendly** - AI tools understand intent better

**Default to PascalCase unless you have a specific reason to use another style.**
