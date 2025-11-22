# excel_worksheet Tool

**Related tools**:
- excel_range - For data operations on worksheet cells
- excel_table - For structured tables on worksheets
- excel_powerquery - For loading external data to worksheets

**Actions**: list, create, rename, copy, delete, move, copy-to-workbook, move-to-workbook, set-tab-color, get-tab-color, clear-tab-color, hide, very-hide, show, get-visibility, set-visibility

**When to use excel_worksheet**:
- Sheet lifecycle (create, delete, rename, copy, move)
- Copy/move sheets within same workbook
- Copy/move sheets between different workbooks
- Sheet visibility (hide, show)
- Sheet tab colors
- Use excel_range for data operations
- Use excel_powerquery for external data loading

## Cross-Workbook Operations

**When to use copy-to-workbook vs copy**:
- `copy`: Duplicate sheet WITHIN same workbook (single sessionId)
- `copy-to-workbook`: Copy sheet TO DIFFERENT workbook (requires TWO sessionIds)
- `move-to-workbook`: Move sheet TO DIFFERENT workbook (requires TWO sessionIds)

**Cross-workbook workflow**:
1. Open source file: `excel_file(action: 'open', excelPath: 'source.xlsx')` → get sessionId1
2. Open target file: `excel_file(action: 'open', excelPath: 'target.xlsx')` → get sessionId2
3. Copy/move sheet: `excel_worksheet(action: 'copy-to-workbook', sessionId: sessionId1, targetSessionId: sessionId2, sheetName: 'Sales', targetName: 'Q1_Sales')`

**Parameters for cross-workbook**:
- `sessionId`: Source workbook session
- `targetSessionId`: Target workbook session
- `sheetName`: Sheet to copy/move from source
- `targetName` (optional): Rename during copy/move
- `beforeSheet` OR `afterSheet` (optional): Position in target workbook

**Example - Copy sheet to consolidation workbook**:
```
User: "Copy the Sales sheet from Q1.xlsx to Annual_Report.xlsx"

1. Open Q1.xlsx → sessionId: "abc123"
2. Open Annual_Report.xlsx → sessionId: "def456"
3. excel_worksheet(
     action: 'copy-to-workbook',
     sessionId: 'abc123',
     targetSessionId: 'def456',
     sheetName: 'Sales',
     targetName: 'Q1_Sales'
   )
```

**Server-specific behavior**:
- Cannot delete last remaining worksheet (Excel limitation)
- Cannot delete active worksheet while viewing in Excel UI
- very-hide: Hidden from UI and VBA (stronger than hide)

**Action disambiguation**:
- create: Add new blank worksheet
- copy: Duplicate existing worksheet WITHIN same workbook
- move: Reposition existing worksheet WITHIN same workbook
- copy-to-workbook: Copy sheet TO DIFFERENT workbook (requires targetSessionId)
- move-to-workbook: Move sheet TO DIFFERENT workbook (requires targetSessionId)
- hide: Hide from UI tabs (visible in VBA)
- very-hide: Hide from UI and VBA (stronger protection)
- show: Make visible in UI tabs

**Common mistakes**:
- Trying to delete last worksheet → Excel requires at least one sheet
- Not checking if sheet exists before operations → Use list first
- Using copy when files are different → Use copy-to-workbook instead
- Forgetting to open both files before cross-workbook operation → Both need sessionIds
