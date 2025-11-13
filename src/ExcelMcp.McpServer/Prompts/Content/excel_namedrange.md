# excel_namedrange Tool

**Related tools**:
- excel_batch - Use for 2+ named range operations (75-90% faster)
- excel_range - For reading/writing data in named ranges (use sheetName="")
- excel_powerquery - Named ranges can be used as Power Query parameters

**Actions**: list, get, set, create, update, delete

**When to use excel_namedrange**:
- Named ranges as configuration parameters
- Reusable cell references
- Settings, thresholds, dates
- Use excel_range for data operations
- Use excel_namedrange interchangeably (same tool)

**Server-specific behavior**:
- Parameters are named ranges pointing to single cells
- Absolute references recommended: =Sheet1!$A$1
- Parameters accessible across entire workbook

**Action disambiguation**:
- create: Add single named range parameter
- get: Retrieve parameter value
- set: Update parameter value
- update: Change parameter cell reference

**Common mistakes**:
- Missing = prefix in references → Must be =Sheet1!$A$1 not Sheet1!$A$1
- Relative references → Use absolute ($A$1) for parameters

**Workflow optimization**:
- Multiple parameters? Use excel_batch with multiple create calls
- Common pattern: create parameters → use in formulas/queries

