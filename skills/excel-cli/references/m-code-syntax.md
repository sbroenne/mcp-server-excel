# M Code Syntax Reference

## Column/Field Name Quoting (CRITICAL)

M code requires `#"..."` quoting for identifiers with hyphens, spaces, or special characters:

| Column Name | Syntax | Notes |
|-------------|--------|-------|
| `Amount` | `[Amount]` | Simple alphanumeric names work without quotes |
| `Non-Recurring` | `[#"Non-Recurring"]` | **Hyphen requires quoting** — without it, M parses as subtraction! |
| `List Price (USD)` | `[#"List Price (USD)"]` | Spaces/parens require quoting |

**Common mistake:** `[Non-Recurring]` parses as `[Non] - [Recurring]` (subtraction!) and fails with cryptic "The name 'X' wasn't recognized" errors.

**Rule:** If a column name contains anything other than letters, numbers, and underscores, use `[#"Column Name"]` syntax.

## Reading Named Ranges (parameters)

```m
Excel.CurrentWorkbook(){[Name = "Param_Name"]}[Content]{0}[Column1]
```

## Query Chaining

Reference other queries by name directly: `Source = OtherQueryName`
