using Spectre.Console;
using System.Text;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet management commands implementation
/// </summary>
public class SheetCommands : ISheetCommands
{
    public int List(string[] args)
    {
        if (!ValidateArgs(args, 2, "sheet-list <file.xlsx>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[bold]Worksheets in:[/] {Path.GetFileName(args[1])}\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            var sheets = new List<(string Name, int Index, bool Visible)>();

            try
            {
                dynamic sheetsCollection = workbook.Worksheets;
                int count = sheetsCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic sheet = sheetsCollection.Item(i);
                    string name = sheet.Name;
                    int visible = sheet.Visible;
                    sheets.Add((name, i, visible == -1)); // -1 = xlSheetVisible
                }
            }
            catch { }

            if (sheets.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]#[/]");
                table.AddColumn("[bold]Sheet Name[/]");
                table.AddColumn("[bold]Visible[/]");

                foreach (var (name, index, visible) in sheets)
                {
                    table.AddRow(
                        $"[dim]{index}[/]",
                        $"[cyan]{name.EscapeMarkup()}[/]",
                        visible ? "[green]Yes[/]" : "[dim]No[/]"
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine($"[bold]Total:[/] {sheets.Count} worksheets");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No worksheets found[/]");
            }

            return 0;
        });
    }

    public int Read(string[] args)
    {
        if (!ValidateArgs(args, 3, "sheet-read <file.xlsx> <sheet-name> [range]")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            AnsiConsole.MarkupLine($"[yellow]Working Directory:[/] {Environment.CurrentDirectory}");
            AnsiConsole.MarkupLine($"[yellow]Full Path Expected:[/] {Path.GetFullPath(args[1])}");
            return 1;
        }

        var sheetName = args[2];
        var range = args.Length > 3 ? args[3] : null;

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName.EscapeMarkup()}' not found");
                    
                    // Show available sheets for coding agent context
                    try
                    {
                        dynamic sheetsCollection = workbook.Worksheets;
                        int sheetCount = sheetsCollection.Count;
                        
                        if (sheetCount > 0)
                        {
                            AnsiConsole.MarkupLine($"[yellow]Available sheets in {Path.GetFileName(args[1])}:[/]");
                            
                            var availableSheets = new List<string>();
                            for (int i = 1; i <= sheetCount; i++)
                            {
                                try
                                {
                                    dynamic ws = sheetsCollection.Item(i);
                                    string name = ws.Name;
                                    bool visible = ws.Visible == -1;
                                    availableSheets.Add(name);
                                    
                                    string visibilityIcon = visible ? "👁" : "🔒";
                                    AnsiConsole.MarkupLine($"  [cyan]{i}.[/] {name.EscapeMarkup()} {visibilityIcon}");
                                }
                                catch (Exception ex)
                                {
                                    AnsiConsole.MarkupLine($"  [red]{i}.[/] <Error accessing sheet: {ex.Message.EscapeMarkup()}>");
                                }
                            }
                            
                            // Suggest closest match
                            var closestMatch = FindClosestSheetMatch(sheetName, availableSheets);
                            if (!string.IsNullOrEmpty(closestMatch))
                            {
                                AnsiConsole.MarkupLine($"[yellow]Did you mean:[/] [cyan]{closestMatch}[/]");
                                AnsiConsole.MarkupLine($"[dim]Command suggestion:[/] [cyan]ExcelCLI sheet-read \"{args[1]}\" \"{closestMatch}\"{(range != null ? $" \"{range}\"" : "")}[/]");
                            }
                        }
                        else
                        {
                            AnsiConsole.MarkupLine("[red]No worksheets found in workbook[/]");
                        }
                    }
                    catch (Exception listEx)
                    {
                        AnsiConsole.MarkupLine($"[red]Error listing sheets:[/] {listEx.Message.EscapeMarkup()}");
                    }
                    
                    return 1;
                }

                // Validate and process range
                dynamic rangeObj;
                string actualRange;
                
                try
                {
                    if (range != null)
                    {
                        rangeObj = sheet.Range(range);
                        actualRange = range;
                    }
                    else
                    {
                        rangeObj = sheet.UsedRange;
                        if (rangeObj == null)
                        {
                            AnsiConsole.MarkupLine($"[yellow]Sheet '{sheetName.EscapeMarkup()}' appears to be empty (no used range)[/]");
                            AnsiConsole.MarkupLine("[dim]Try adding data to the sheet first[/]");
                            return 0;
                        }
                        actualRange = rangeObj.Address;
                    }
                }
                catch (Exception rangeEx)
                {
                    AnsiConsole.MarkupLine($"[red]Error accessing range '[cyan]{range ?? "UsedRange"}[/]':[/] {rangeEx.Message.EscapeMarkup()}");
                    
                    // Provide guidance for range format
                    if (range != null)
                    {
                        AnsiConsole.MarkupLine("[yellow]Range format examples:[/]");
                        AnsiConsole.MarkupLine("  • [cyan]A1[/] (single cell)");
                        AnsiConsole.MarkupLine("  • [cyan]A1:D10[/] (rectangular range)");
                        AnsiConsole.MarkupLine("  • [cyan]A:A[/] (entire column)");
                        AnsiConsole.MarkupLine("  • [cyan]1:1[/] (entire row)");
                    }
                    return 1;
                }

                object? values = rangeObj.Value;

                if (values == null)
                {
                    AnsiConsole.MarkupLine($"[yellow]No data found in range '{actualRange.EscapeMarkup()}'[/]");
                    return 0;
                }

                AnsiConsole.MarkupLine($"[bold]Reading from:[/] [cyan]{sheetName.EscapeMarkup()}[/] range [cyan]{actualRange.EscapeMarkup()}[/]");
                AnsiConsole.WriteLine();

                // Display data in table
                var table = new Table();
                table.Border(TableBorder.Rounded);

                // Handle single cell
                if (values is not Array)
                {
                    table.AddColumn("Value");
                    table.AddColumn("Type");
                    
                    string cellValue = values?.ToString() ?? "";
                    string valueType = values?.GetType().Name ?? "null";
                    
                    table.AddRow(cellValue.EscapeMarkup(), valueType);
                    AnsiConsole.Write(table);
                    
                    AnsiConsole.MarkupLine($"[dim]Single cell value, type: {valueType}[/]");
                    return 0;
                }

                // Handle array (2D)
                var array = values as object[,];
                if (array == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Unable to read data as array. Data type: {values.GetType().Name}");
                    return 1;
                }

                int rows = array.GetLength(0);
                int cols = array.GetLength(1);

                AnsiConsole.MarkupLine($"[dim]Data dimensions: {rows} rows × {cols} columns[/]");

                // Add columns (use first row as headers if looks like headers, else Col1, Col2, etc.)
                for (int col = 1; col <= cols; col++)
                {
                    var headerVal = array[1, col]?.ToString() ?? $"Col{col}";
                    table.AddColumn($"[bold]{headerVal.EscapeMarkup()}[/]");
                }

                // Add rows (skip first row if using as headers)
                int dataRows = 0;
                int startRow = rows > 1 ? 2 : 1; // Skip first row if multiple rows (assume headers)
                
                for (int row = startRow; row <= rows; row++)
                {
                    var rowData = new List<string>();
                    for (int col = 1; col <= cols; col++)
                    {
                        var cellValue = array[row, col];
                        string displayValue = cellValue?.ToString() ?? "";
                        
                        // Truncate very long values for display
                        if (displayValue.Length > 100)
                        {
                            displayValue = displayValue[..97] + "...";
                        }
                        
                        rowData.Add(displayValue.EscapeMarkup());
                    }
                    table.AddRow(rowData.ToArray());
                    dataRows++;
                    
                    // Limit display for very large datasets
                    if (dataRows >= 50)
                    {
                        table.AddRow(Enumerable.Repeat($"[dim]... ({rows - row} more rows)[/]", cols).ToArray());
                        break;
                    }
                }

                AnsiConsole.Write(table);
                AnsiConsole.WriteLine();
                
                if (rows > 1)
                {
                    AnsiConsole.MarkupLine($"[dim]Displayed {Math.Min(dataRows, rows - 1)} data rows (excluding header)[/]");
                }
                else
                {
                    AnsiConsole.MarkupLine($"[dim]Displayed {dataRows} rows[/]");
                }

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error reading sheet data:[/] {ex.Message.EscapeMarkup()}");
                
                // Provide additional context for coding agents
                ExcelDiagnostics.ReportOperationContext("sheet-read", args[1],
                    ("Sheet", sheetName),
                    ("Range", range ?? "UsedRange"),
                    ("Error Type", ex.GetType().Name));
                    
                return 1;
            }
        });
    }

    /// <summary>
    /// Finds the closest matching sheet name
    /// </summary>
    private static string? FindClosestSheetMatch(string target, List<string> candidates)
    {
        if (candidates.Count == 0) return null;
        
        // First try exact case-insensitive match
        var exactMatch = candidates.FirstOrDefault(c => 
            string.Equals(c, target, StringComparison.OrdinalIgnoreCase));
        if (exactMatch != null) return exactMatch;
        
        // Then try substring match
        var substringMatch = candidates.FirstOrDefault(c => 
            c.Contains(target, StringComparison.OrdinalIgnoreCase) || 
            target.Contains(c, StringComparison.OrdinalIgnoreCase));
        if (substringMatch != null) return substringMatch;
        
        // Finally use Levenshtein distance
        int minDistance = int.MaxValue;
        string? bestMatch = null;
        
        foreach (var candidate in candidates)
        {
            int distance = ComputeLevenshteinDistance(target.ToLowerInvariant(), candidate.ToLowerInvariant());
            if (distance < minDistance && distance <= Math.Max(target.Length, candidate.Length) / 2)
            {
                minDistance = distance;
                bestMatch = candidate;
            }
        }
        
        return bestMatch;
    }
    
    /// <summary>
    /// Computes Levenshtein distance between two strings
    /// </summary>
    private static int ComputeLevenshteinDistance(string s1, string s2)
    {
        int[,] d = new int[s1.Length + 1, s2.Length + 1];
        
        for (int i = 0; i <= s1.Length; i++)
            d[i, 0] = i;
        for (int j = 0; j <= s2.Length; j++)
            d[0, j] = j;
            
        for (int i = 1; i <= s1.Length; i++)
        {
            for (int j = 1; j <= s2.Length; j++)
            {
                int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        }
        
        return d[s1.Length, s2.Length];
    }

    public async Task<int> Write(string[] args)
    {
        if (!ValidateArgs(args, 4, "sheet-write <file.xlsx> <sheet-name> <data.csv>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }
        if (!File.Exists(args[3]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] CSV file not found: {args[3]}");
            return 1;
        }

        var sheetName = args[2];
        var csvFile = args[3];

        // Read CSV
        var lines = await File.ReadAllLinesAsync(csvFile);
        if (lines.Length == 0)
        {
            AnsiConsole.MarkupLine("[yellow]CSV file is empty[/]");
            return 1;
        }

        var data = new List<string[]>();
        foreach (var line in lines)
        {
            // Simple CSV parsing (doesn't handle quoted commas)
            data.Add(line.Split(','));
        }

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            dynamic? sheet = FindSheet(workbook, sheetName);
            if (sheet == null)
            {
                // Create new sheet
                dynamic sheetsCollection = workbook.Worksheets;
                sheet = sheetsCollection.Add();
                sheet.Name = sheetName;
                AnsiConsole.MarkupLine($"[yellow]Created new sheet '{sheetName}'[/]");
            }

            // Clear existing data
            dynamic usedRange = sheet.UsedRange;
            try { usedRange.Clear(); } catch { }

            // Write data
            int rows = data.Count;
            int cols = data[0].Length;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (j < data[i].Length)
                    {
                        dynamic cell = sheet.Cells[i + 1, j + 1];
                        cell.Value = data[i][j];
                    }
                }
            }

            workbook.Save();
            AnsiConsole.MarkupLine($"[green]✓[/] Wrote {rows} rows × {cols} columns to sheet '{sheetName}'");
            return 0;
        });
    }

    public int Copy(string[] args)
    {
        if (!ValidateArgs(args, 4, "sheet-copy <file.xlsx> <source-sheet> <new-sheet>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sourceSheet = args[2];
        var newSheet = args[3];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            dynamic? sheet = FindSheet(workbook, sourceSheet);
            if (sheet == null)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sourceSheet}' not found");
                return 1;
            }

            // Check if target already exists
            if (FindSheet(workbook, newSheet) != null)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{newSheet}' already exists");
                return 1;
            }

            // Copy sheet
            sheet.Copy(After: workbook.Worksheets[workbook.Worksheets.Count]);
            dynamic copiedSheet = workbook.Worksheets[workbook.Worksheets.Count];
            copiedSheet.Name = newSheet;

            workbook.Save();
            AnsiConsole.MarkupLine($"[green]✓[/] Copied sheet '{sourceSheet}' to '{newSheet}'");
            return 0;
        });
    }

    public int Delete(string[] args)
    {
        if (!ValidateArgs(args, 3, "sheet-delete <file.xlsx> <sheet-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            dynamic? sheet = FindSheet(workbook, sheetName);
            if (sheet == null)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                return 1;
            }

            // Prevent deleting the last sheet
            if (workbook.Worksheets.Count == 1)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Cannot delete the last worksheet");
                return 1;
            }

            sheet.Delete();
            workbook.Save();
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted sheet '{sheetName}'");
            return 0;
        });
    }

    public int Create(string[] args)
    {
        if (!ValidateArgs(args, 3, "sheet-create <file.xlsx> <sheet-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                // Check if sheet already exists
                dynamic? existingSheet = FindSheet(workbook, sheetName);
                if (existingSheet != null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' already exists");
                    return 1;
                }

                // Add new worksheet
                dynamic sheets = workbook.Worksheets;
                dynamic newSheet = sheets.Add();
                newSheet.Name = sheetName;

                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Created sheet '{sheetName}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Rename(string[] args)
    {
        if (!ValidateArgs(args, 4, "sheet-rename <file.xlsx> <old-name> <new-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var oldName = args[2];
        var newName = args[3];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, oldName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{oldName}' not found");
                    return 1;
                }

                // Check if new name already exists
                dynamic? existingSheet = FindSheet(workbook, newName);
                if (existingSheet != null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{newName}' already exists");
                    return 1;
                }

                sheet.Name = newName;
                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Renamed sheet '{oldName}' to '{newName}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Clear(string[] args)
    {
        if (!ValidateArgs(args, 3, "sheet-clear <file.xlsx> <sheet-name> (range)")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];
        var range = args.Length > 3 ? args[3] : "A:XFD"; // Clear entire sheet if no range specified

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                    return 1;
                }

                dynamic targetRange = sheet.Range[range];
                targetRange.Clear();

                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Cleared range '{range}' in sheet '{sheetName}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Append(string[] args)
    {
        if (!ValidateArgs(args, 4, "sheet-append <file.xlsx> <sheet-name> <data-file.csv>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }
        if (!File.Exists(args[3]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Data file not found: {args[3]}");
            return 1;
        }

        var sheetName = args[2];
        var dataFile = args[3];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                    return 1;
                }

                // Read CSV data
                var lines = File.ReadAllLines(dataFile);
                if (lines.Length == 0)
                {
                    AnsiConsole.MarkupLine("[yellow]Warning:[/] Data file is empty");
                    return 0;
                }

                // Find the last used row
                dynamic usedRange = sheet.UsedRange;
                int lastRow = usedRange != null ? usedRange.Rows.Count : 0;
                int startRow = lastRow + 1;

                // Parse CSV and write data
                for (int i = 0; i < lines.Length; i++)
                {
                    var values = lines[i].Split(',');
                    for (int j = 0; j < values.Length; j++)
                    {
                        dynamic cell = sheet.Cells[startRow + i, j + 1];
                        cell.Value2 = values[j].Trim('"');
                    }
                }

                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Appended {lines.Length} rows to sheet '{sheetName}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }
}
