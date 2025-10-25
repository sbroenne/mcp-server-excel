using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet management commands implementation
/// </summary>
public class SheetCommands : ISheetCommands
{
    /// <inheritdoc />
    public WorksheetListResult List(string filePath)
    {
        if (!File.Exists(filePath))
            return new WorksheetListResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath };

        var result = new WorksheetListResult { FilePath = filePath };
        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        result.Worksheets.Add(new WorksheetInfo { Name = sheet.Name, Index = i });
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheet);
                    }
                }
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public WorksheetDataResult Read(string filePath, string sheetName, string range)
    {
        if (!File.Exists(filePath))
            return new WorksheetDataResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath };

        var result = new WorksheetDataResult { FilePath = filePath, SheetName = sheetName, Range = range };
        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? rangeObj = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }

                rangeObj = sheet.Range[range];
                object[,] values = rangeObj.Value2;
                if (values != null)
                {
                    int rows = values.GetLength(0), cols = values.GetLength(1);
                    for (int r = 1; r <= rows; r++)
                    {
                        var row = new List<object?>();
                        for (int c = 1; c <= cols; c++) row.Add(values[r, c]);
                        result.Data.Add(row);
                    }
                }
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref rangeObj);
                ComUtilities.Release(ref sheet);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Write(string filePath, string sheetName, string csvData)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "write" };

        var result = new OperationResult { FilePath = filePath, Action = "write" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell1 = null;
            dynamic? cell2 = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }

                var data = ParseCsv(csvData);
                if (data.Count == 0) { result.Success = false; result.ErrorMessage = "No data to write"; return 1; }

                int rows = data.Count, cols = data[0].Count;
                object[,] arr = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = data[r][c];

                cell1 = sheet.Cells[1, 1];
                cell2 = sheet.Cells[rows, cols];
                range = sheet.Range[cell1, cell2];
                range.Value2 = arr;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref cell2);
                ComUtilities.Release(ref cell1);
                ComUtilities.Release(ref sheet);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Create(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "create-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "create-sheet" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheets = null;
            dynamic? newSheet = null;
            try
            {
                sheets = workbook.Worksheets;
                newSheet = sheets.Add();
                newSheet.Name = sheetName;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Rename(string filePath, string oldName, string newName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "rename-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "rename-sheet" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, oldName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{oldName}' not found"; return 1; }
                sheet.Name = newName;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Copy(string filePath, string sourceName, string targetName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "copy-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "copy-sheet" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sourceSheet = null;
            dynamic? sheets = null;
            dynamic? lastSheet = null;
            dynamic? copiedSheet = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(workbook, sourceName);
                if (sourceSheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sourceName}' not found"; return 1; }
                sheets = workbook.Worksheets;
                lastSheet = sheets.Item(sheets.Count);
                sourceSheet.Copy(After: lastSheet);
                copiedSheet = sheets.Item(sheets.Count);
                copiedSheet.Name = targetName;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref copiedSheet);
                ComUtilities.Release(ref lastSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref sourceSheet);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Delete(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "delete-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "delete-sheet" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }
                sheet.Delete();
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Clear(string filePath, string sheetName, string range)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "clear" };

        var result = new OperationResult { FilePath = filePath, Action = "clear" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? rangeObj = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }
                rangeObj = sheet.Range[range];
                rangeObj.Clear();
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref rangeObj);
                ComUtilities.Release(ref sheet);
            }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Append(string filePath, string sheetName, string csvData)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "append" };

        var result = new OperationResult { FilePath = filePath, Action = "append" };
        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? usedRange = null;
            dynamic? rows = null;
            dynamic? cell1 = null;
            dynamic? cell2 = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }

                usedRange = sheet.UsedRange;
                rows = usedRange.Rows;
                int lastRow = rows.Count;

                var data = ParseCsv(csvData);
                if (data.Count == 0) { result.Success = false; result.ErrorMessage = "No data to append"; return 1; }

                int startRow = lastRow + 1, numRows = data.Count, cols = data[0].Count;
                object[,] arr = new object[numRows, cols];
                for (int r = 0; r < numRows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = data[r][c];

                cell1 = sheet.Cells[startRow, 1];
                cell2 = sheet.Cells[startRow + numRows - 1, cols];
                range = sheet.Range[cell1, cell2];
                range.Value2 = arr;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref cell2);
                ComUtilities.Release(ref cell1);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref usedRange);
                ComUtilities.Release(ref sheet);
            }
        });
        return result;
    }

    private static List<List<string>> ParseCsv(string csvData)
    {
        var result = new List<List<string>>();
        var lines = csvData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var row = new List<string>();
            var fields = line.Split(',');
            foreach (var field in fields)
                row.Add(field.Trim().Trim('"'));
            result.Add(row);
        }
        return result;
    }

    /// <inheritdoc />
    public OperationResult Protect(string filePath, string sheetName, string? password = null, 
        bool allowFormatCells = false, bool allowFormatColumns = false, bool allowFormatRows = false,
        bool allowInsertColumns = false, bool allowInsertRows = false, bool allowDeleteColumns = false,
        bool allowDeleteRows = false, bool allowSort = false, bool allowFilter = false)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "protect"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "protect"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                // Protect the sheet with specified permissions
                sheet.Protect(
                    Password: password,
                    DrawingObjects: true,
                    Contents: true,
                    Scenarios: true,
                    AllowFormattingCells: allowFormatCells,
                    AllowFormattingColumns: allowFormatColumns,
                    AllowFormattingRows: allowFormatRows,
                    AllowInsertingColumns: allowInsertColumns,
                    AllowInsertingRows: allowInsertRows,
                    AllowDeletingColumns: allowDeleteColumns,
                    AllowDeletingRows: allowDeleteRows,
                    AllowSorting: allowSort,
                    AllowFiltering: allowFilter
                );

                result.Success = true;
                result.WorkflowHint = $"Sheet '{sheetName}' protected" + (password != null ? " with password" : "");
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Unprotect(string filePath, string sheetName, string? password = null)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "unprotect"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "unprotect"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                // Unprotect the sheet
                if (password != null)
                {
                    sheet.Unprotect(password);
                }
                else
                {
                    sheet.Unprotect();
                }

                result.Success = true;
                result.WorkflowHint = $"Sheet '{sheetName}' unprotected";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public SheetProtectionResult GetProtectionStatus(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
        {
            return new SheetProtectionResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                SheetName = sheetName
            };
        }

        var result = new SheetProtectionResult
        {
            FilePath = filePath,
            SheetName = sheetName
        };

        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? protection = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                result.IsProtected = sheet.ProtectContents;
                
                if (result.IsProtected)
                {
                    protection = sheet.Protection;
                    result.AllowFormatCells = protection.AllowFormattingCells;
                    result.AllowFormatColumns = protection.AllowFormattingColumns;
                    result.AllowFormatRows = protection.AllowFormattingRows;
                    result.AllowInsertColumns = protection.AllowInsertingColumns;
                    result.AllowInsertRows = protection.AllowInsertingRows;
                    result.AllowDeleteColumns = protection.AllowDeletingColumns;
                    result.AllowDeleteRows = protection.AllowDeletingRows;
                    result.AllowSort = protection.AllowSorting;
                    result.AllowFilter = protection.AllowFiltering;
                }

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref protection);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }
}
