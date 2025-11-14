# excel_namedrange Tool

**Related tools**:
- excel_range - For reading/writing data in named ranges (use sheetName="")
- excel_powerquery - Named ranges can be used as Power Query parameters

**Actions**: list, get, set, create, update, delete, create-bulk

**When to use excel_namedrange**:
- Named ranges as configuration parameters
- Reusable cell references
- Settings, thresholds, dates
- Use excel_range for data operations
- Use excel_namedrange interchangeably (same tool)

**Server-specific behavior**:
- Parameters are named ranges pointing to single cells
- create-bulk: Efficient multi-parameter creation (one call vs many)
- Absolute references recommended: =Sheet1!$A$1
- Parameters accessible across entire workbook

**Action disambiguation**:
- create: Add single named range parameter
- create-bulk: Add multiple parameters in one call (90% faster)
- get: Retrieve parameter value
- set: Update parameter value
- update: Change parameter cell reference

**Common mistakes**:
- Creating parameters one-by-one → Use create-bulk for 2+ parameters
- Missing = prefix in references → Must be =Sheet1!$A$1 not Sheet1!$A$1
- Relative references → Use absolute ($A$1) for parameters

**Workflow optimization**:
- Multiple parameters? Use create-bulk action
