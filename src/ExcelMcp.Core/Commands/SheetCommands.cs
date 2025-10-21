using Sbroenne.ExcelMcp.Core.Models;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

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
        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic sheets = workbook.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    result.Worksheets.Add(new WorksheetInfo { Name = sheet.Name, Index = i });
                }
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public WorksheetDataResult Read(string filePath, string sheetName, string range)
    {
        if (!File.Exists(filePath))
            return new WorksheetDataResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath };

        var result = new WorksheetDataResult { FilePath = filePath, SheetName = sheetName, Range = range };
        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }

                dynamic rangeObj = sheet.Range[range];
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
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Write(string filePath, string sheetName, string csvData)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "write" };

        var result = new OperationResult { FilePath = filePath, Action = "write" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }

                var data = ParseCsv(csvData);
                if (data.Count == 0) { result.Success = false; result.ErrorMessage = "No data to write"; return 1; }

                int rows = data.Count, cols = data[0].Count;
                object[,] arr = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = data[r][c];

                dynamic range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rows, cols]];
                range.Value2 = arr;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Create(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "create-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "create-sheet" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic sheets = workbook.Worksheets;
                dynamic newSheet = sheets.Add();
                newSheet.Name = sheetName;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Rename(string filePath, string oldName, string newName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "rename-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "rename-sheet" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, oldName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{oldName}' not found"; return 1; }
                sheet.Name = newName;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Copy(string filePath, string sourceName, string targetName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "copy-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "copy-sheet" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic? sourceSheet = FindSheet(workbook, sourceName);
                if (sourceSheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sourceName}' not found"; return 1; }
                sourceSheet.Copy(After: workbook.Worksheets.Item(workbook.Worksheets.Count));
                dynamic copiedSheet = workbook.Worksheets.Item(workbook.Worksheets.Count);
                copiedSheet.Name = targetName;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Delete(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "delete-sheet" };

        var result = new OperationResult { FilePath = filePath, Action = "delete-sheet" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }
                sheet.Delete();
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Clear(string filePath, string sheetName, string range)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "clear" };

        var result = new OperationResult { FilePath = filePath, Action = "clear" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }
                dynamic rangeObj = sheet.Range[range];
                rangeObj.Clear();
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
        });
        return result;
    }

    /// <inheritdoc />
    public OperationResult Append(string filePath, string sheetName, string csvData)
    {
        if (!File.Exists(filePath))
            return new OperationResult { Success = false, ErrorMessage = $"File not found: {filePath}", FilePath = filePath, Action = "append" };

        var result = new OperationResult { FilePath = filePath, Action = "append" };
        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null) { result.Success = false; result.ErrorMessage = $"Sheet '{sheetName}' not found"; return 1; }

                dynamic usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;

                var data = ParseCsv(csvData);
                if (data.Count == 0) { result.Success = false; result.ErrorMessage = "No data to append"; return 1; }

                int startRow = lastRow + 1, rows = data.Count, cols = data[0].Count;
                object[,] arr = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        arr[r, c] = data[r][c];

                dynamic range = sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[startRow + rows - 1, cols]];
                range.Value2 = arr;
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex) { result.Success = false; result.ErrorMessage = ex.Message; return 1; }
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
}
