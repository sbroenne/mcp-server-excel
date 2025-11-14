# excel_vba Tool

**Related tools**:
- excel_file - Create .xlsm files (macro-enabled workbooks)
- excel_range - For reading/writing data from VBA procedures

**Actions**: list, view, import, export, update, delete, run

**When to use excel_vba**:
- VBA macro management
- Import/export VBA modules for version control
- Run existing macros
- Use excel_range for data operations
- Requires .xlsm files (macro-enabled)

**Server-specific behavior**:
- Requires macro-enabled workbook (.xlsm)
- VBA trust settings must allow programmatic access
- Modules stored as .bas or .cls files
- run action executes VBA Sub procedures

**Action disambiguation**:
- list: Show all VBA modules in workbook
- view: Get VBA code for specific module
- import: Load VBA code from .bas/.cls file
- export: Save VBA module to file (version control)
- run: Execute VBA Sub procedure
- update: Replace module code

**Common mistakes**:
- Using .xlsx instead of .xlsm → VBA requires macro-enabled files
- VBA trust not enabled → Check security settings
- Running Function instead of Sub → Use run for Sub only

**Workflow optimization**:
- Version control: export → Git → import on other machine
