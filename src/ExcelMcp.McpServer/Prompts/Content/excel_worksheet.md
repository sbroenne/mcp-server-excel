# excel_worksheet Tool

**Related tools**:
- excel_batch - Use for 2+ worksheet operations (75-90% faster)
- excel_range - For data operations on worksheet cells
- excel_table - For structured tables on worksheets
- excel_powerquery - For loading external data to worksheets

**Actions**: list, create, rename, copy, delete, set-tab-color, get-tab-color, clear-tab-color, hide, very-hide, show, get-visibility, set-visibility

**When to use excel_worksheet**:
- Sheet lifecycle (create, delete, rename, copy)
- Sheet visibility (hide, show)
- Sheet tab colors
- Use excel_range for data operations
- Use excel_powerquery for external data loading

**Server-specific behavior**:
- Cannot delete last remaining worksheet (Excel limitation)
- Cannot delete active worksheet while viewing in Excel UI
- very-hide: Hidden from UI and VBA (stronger than hide)
- Batch mode recommended for creating multiple sheets

**Action disambiguation**:
- create: Add new blank worksheet
- copy: Duplicate existing worksheet
- hide: Hide from UI tabs (visible in VBA)
- very-hide: Hide from UI and VBA (stronger protection)
- show: Make visible in UI tabs

**Common mistakes**:
- Trying to delete last worksheet → Excel requires at least one sheet
- Creating sheets one-by-one → Use batch mode for multiple sheets
- Not checking if sheet exists before operations → Use list first

**Workflow optimization**:
- Creating multiple sheets? Use begin_excel_batch
- After creating sheets: Use excel_range to populate data
- Use set-tab-color to organize sheets visually
