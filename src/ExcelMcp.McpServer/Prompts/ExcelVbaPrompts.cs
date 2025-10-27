using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts for VBA development patterns and best practices.
/// </summary>
[McpServerPromptType]
public static class ExcelVbaPrompts
{
    /// <summary>
    /// VBA development guide with error handling and best practices.
    /// </summary>
    [McpServerPrompt(Name = "excel_vba_guide")]
    [Description("VBA development patterns, error handling, and automation best practices")]
    public static ChatMessage VbaDevelopmentGuide()
    {
        return new ChatMessage(ChatRole.User, @"# VBA Development Guide

## VBA Trust Setup (One-Time)

Before using VBA commands, enable VBA trust in Excel:
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings
3. ✓ Check ""Trust access to the VBA project object model""

## VBA Module Management

### List All Modules
```bash
excel_vba list --file report.xlsm
```

### Export Module for Version Control
```bash
excel_vba export --file report.xlsm --module DataProcessor --output vba/processor.vba
```

### Import/Update Module
```bash
excel_vba import --file report.xlsm --module DataProcessor --source vba/processor.vba
```

### Run Macro
```bash
excel_vba run --file report.xlsm --macro ProcessData
excel_vba run --file report.xlsm --macro CalculateTotal --args Sheet1 A1:C10
```

## VBA Best Practices

### 1. Error Handling (Always Required)
```vba
Sub ProcessData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(""Data"")
    
    ' Your logic here
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Error "" & Err.Number & "": "" & Err.Description
    Err.Clear
End Sub
```

### 2. Use Option Explicit (Prevent Typos)
```vba
Option Explicit  ' At top of every module

Sub MyMacro()
    Dim rowCount As Long  ' Must declare all variables
    rowCount = ActiveSheet.UsedRange.Rows.Count
End Sub
```

### 3. Release Object References
```vba
Sub ProcessWorkbook()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Open(""C:\data.xlsx"")
    Set ws = wb.Sheets(1)
    
    ' Process data
    
    ' Cleanup
    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing
End Sub
```

### 4. Avoid Select/Activate (Faster)
```vba
' ❌ SLOW - Uses Select
Range(""A1"").Select
Selection.Value = ""Hello""

' ✅ FAST - Direct assignment
Range(""A1"").Value = ""Hello""
```

### 5. Turn Off Screen Updating (Large Operations)
```vba
Sub BulkUpdate()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Your intensive operations
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

## Common VBA Patterns

### Pattern 1: Loop Through Range
```vba
Sub ProcessRange()
    Dim ws As Worksheet
    Dim cell As Range
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Sheets(""Data"")
    Set dataRange = ws.Range(""A2:A100"")
    
    For Each cell In dataRange
        If cell.Value > 100 Then
            cell.Offset(0, 1).Value = ""High""
        End If
    Next cell
End Sub
```

### Pattern 2: Array Processing (Faster)
```vba
Sub ProcessWithArray()
    Dim ws As Worksheet
    Dim dataArray As Variant
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets(""Data"")
    dataArray = ws.Range(""A2:C100"").Value
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        If dataArray(i, 1) > 100 Then
            dataArray(i, 2) = ""High""
        End If
    Next i
    
    ws.Range(""A2:C100"").Value = dataArray
End Sub
```

### Pattern 3: Call Power Query Refresh
```vba
Sub RefreshQueries()
    On Error Resume Next
    
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        If InStr(conn.Name, ""Query"") > 0 Then
            conn.Refresh
        End If
    Next conn
End Sub
```

### Pattern 4: Create Worksheet from Template
```vba
Sub CreateMonthlyReport(reportMonth As String)
    Dim templateWs As Worksheet
    Dim newWs As Worksheet
    
    Set templateWs = ThisWorkbook.Sheets(""Template"")
    Set newWs = templateWs.Copy(After:=ThisWorkbook.Sheets(Sheets.Count))
    
    newWs.Name = reportMonth
    newWs.Range(""A1"").Value = ""Report for "" & reportMonth
End Sub
```

## Version Control Workflow

### 1. Export All Modules
```bash
for module in $(excel_vba list --file report.xlsm | grep Module); do
    excel_vba export --file report.xlsm --module $module --output vba/$module.vba
done
```

### 2. Track in Git
```bash
git add vba/*.vba
git commit -m ""Export VBA modules for version control""
```

### 3. Update from Git
```bash
git pull
excel_vba import --file report.xlsm --module DataProcessor --source vba/DataProcessor.vba
```

## Security Best Practices

1. **Never store passwords** in VBA code
2. **Code sign macros** for distribution
3. **Use early binding** (reference specific libraries)
4. **Validate user input** before processing
5. **Don't auto-run macros** on workbook open (security risk)

## Debugging Tips

1. **Add Debug.Print statements**
2. **Use Immediate Window** (Ctrl+G in VBA editor)
3. **Set breakpoints** in VBA editor
4. **Use Watch Window** for variable inspection
5. **Check Err.Number** in error handlers");
    }

    /// <summary>
    /// Guide for integrating VBA with Power Query and worksheets.
    /// </summary>
    [McpServerPrompt(Name = "excel_vba_integration")]
    [Description("Integrate VBA with Power Query, worksheets, and parameters")]
    public static ChatMessage VbaIntegration()
    {
        return new ChatMessage(ChatRole.User, @"# VBA Integration Patterns

## VBA + Power Query

### Refresh Specific Query
```vba
Sub RefreshSalesQuery()
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        If conn.Name = ""Query - SalesData"" Or conn.Name = ""SalesData"" Then
            conn.Refresh
            Exit For
        End If
    Next conn
End Sub
```

### Refresh Multiple Queries in Order
```vba
Sub RefreshInOrder()
    ' Refresh helper queries first
    RefreshQuery ""DateDimension""
    RefreshQuery ""ProductCatalog""
    
    ' Then main queries
    RefreshQuery ""Sales""
    RefreshQuery ""Inventory""
End Sub

Sub RefreshQuery(queryName As String)
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        If conn.Name = queryName Or conn.Name = ""Query - "" & queryName Then
            conn.Refresh
            Exit Sub
        End If
    Next conn
End Sub
```

## VBA + Named Ranges (Parameters)

### Read Parameter Value
```vba
Sub ReadParameter()
    Dim startDate As Date
    
    On Error Resume Next
    startDate = Range(""StartDate"").Value
    
    If Err.Number <> 0 Then
        MsgBox ""Parameter 'StartDate' not found""
        Exit Sub
    End If
    
    MsgBox ""Start Date: "" & startDate
End Sub
```

### Update Parameter Before Refresh
```vba
Sub UpdateAndRefresh()
    ' Update parameters
    Range(""StartDate"").Value = DateSerial(2024, 1, 1)
    Range(""EndDate"").Value = Date
    
    ' Refresh queries that use parameters
    RefreshQuery ""SalesData""
    
    MsgBox ""Data refreshed for "" & Range(""StartDate"").Value & "" to "" & Range(""EndDate"").Value
End Sub
```

## VBA + Worksheets

### Create Worksheet from Query Results
```vba
Sub CreateSummarySheet()
    Dim sourceWs As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    
    ' Get source data (from Power Query)
    Set sourceWs = ThisWorkbook.Sheets(""SalesData"")
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    
    ' Create summary sheet
    Set summaryWs = ThisWorkbook.Sheets.Add
    summaryWs.Name = ""Summary_"" & Format(Date, ""yyyymmdd"")
    
    ' Copy and summarize data
    sourceWs.Range(""A1:D"" & lastRow).Copy
    summaryWs.Range(""A1"").PasteSpecial xlPasteValues
    
    ' Add pivot table or formulas
    summaryWs.Range(""F2"").Formula = ""=SUMIF(A:A,'Customer','C:C')""
End Sub
```

### Loop Through Query Results
```vba
Sub ProcessQueryResults()
    Dim ws As Worksheet
    Dim dataArray As Variant
    Dim i As Long
    Dim total As Double
    
    Set ws = ThisWorkbook.Sheets(""SalesData"")
    dataArray = ws.Range(""A2:C"" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Value
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        total = total + dataArray(i, 3)  ' Sum column C
    Next i
    
    MsgBox ""Total Sales: "" & Format(total, ""$#,##0.00"")
End Sub
```

## Complete Automation Workflow

### Automated Monthly Report
```vba
Sub GenerateMonthlyReport()
    On Error GoTo ErrorHandler
    
    Dim reportMonth As String
    reportMonth = Format(DateAdd(""m"", -1, Date), ""yyyy-mm"")
    
    ' 1. Update parameters
    Range(""ReportMonth"").Value = reportMonth
    
    ' 2. Refresh all queries
    RefreshAllQueries
    
    ' 3. Create summary sheet
    CreateSummarySheet reportMonth
    
    ' 4. Generate charts
    CreateCharts
    
    ' 5. Save as PDF
    ExportToPDF ""Reports\Monthly_"" & reportMonth & "".pdf""
    
    MsgBox ""Report generated successfully for "" & reportMonth
    Exit Sub
    
ErrorHandler:
    MsgBox ""Error generating report: "" & Err.Description
End Sub

Sub RefreshAllQueries()
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        conn.Refresh
    Next conn
    
    ' Wait for refresh to complete
    Application.CalculateUntilAsyncQueriesDone
End Sub

Sub ExportToPDF(filePath As String)
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard
End Sub
```

## Batch Session + VBA Pattern

For complex workflows, combine batch sessions with VBA:

```typescript
// 1. Begin batch session
const { batchId } = await begin_excel_batch({ filePath: ""report.xlsm"" });

// 2. Update queries via batch
await excel_powerquery({ batchId, action: ""refresh"", queryName: ""Sales"" });
await excel_powerquery({ batchId, action: ""refresh"", queryName: ""Products"" });

// 3. Run VBA to process results
await excel_vba({ 
    batchId, 
    action: ""run"", 
    macroName: ""GenerateSummary"" 
});

// 4. Export final result
await excel_vba({
    batchId,
    action: ""run"",
    macroName: ""ExportToPDF"",
    args: [""Reports/output.pdf""]
});

// 5. Commit
await commit_excel_batch({ batchId, save: true });
```

## Best Practices

1. **Separate concerns** - VBA for UI/automation, Power Query for data
2. **Error handling** - Always use On Error in VBA
3. **Wait for refresh** - Use Application.CalculateUntilAsyncQueriesDone
4. **Version control** - Export VBA modules regularly
5. **Test incrementally** - Test each step before combining");
    }
}
