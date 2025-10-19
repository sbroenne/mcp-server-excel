using Spectre.Console;
using static Sbroenne.ExcelMcp.CLI.ExcelHelper;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Power Query management commands implementation
/// </summary>
public class PowerQueryCommands : IPowerQueryCommands
{
    /// <summary>
    /// Finds the closest matching string using simple Levenshtein distance
    /// </summary>
    private static string? FindClosestMatch(string target, List<string> candidates)
    {
        if (candidates.Count == 0) return null;
        
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
    public int List(string[] args)
    {
        if (!ValidateArgs(args, 2, "pq-list <file.xlsx>")) return 1;
        if (!ValidateExcelFile(args[1])) return 1;

        AnsiConsole.MarkupLine($"[bold]Power Queries in:[/] {Path.GetFileName(args[1])}\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            var queries = new List<(string Name, string Formula)>();

            try
            {
                // Get Power Queries with enhanced error handling
                dynamic queriesCollection = workbook.Queries;
                int count = queriesCollection.Count;
                
                AnsiConsole.MarkupLine($"[dim]Found {count} Power Queries[/]");
                
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";
                        string formula = query.Formula ?? "";
                        
                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;
                        queries.Add((name, preview));
                    }
                    catch (Exception queryEx)
                    {
                        AnsiConsole.MarkupLine($"[yellow]Warning:[/] Error accessing query {i}: {queryEx.Message.EscapeMarkup()}");
                        queries.Add(($"Error Query {i}", $"{queryEx.Message}"));
                    }
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error accessing Power Queries:[/] {ex.Message.EscapeMarkup()}");
                
                // Check if this workbook supports Power Query
                try
                {
                    string fileName = Path.GetFileName(args[1]);
                    string extension = Path.GetExtension(args[1]).ToLowerInvariant();
                    
                    if (extension == ".xls")
                    {
                        AnsiConsole.MarkupLine("[yellow]Note:[/] .xls files don't support Power Query. Use .xlsx or .xlsm");
                    }
                    else
                    {
                        AnsiConsole.MarkupLine("[yellow]This workbook may not have Power Query enabled[/]");
                        AnsiConsole.MarkupLine("[dim]Try opening the file in Excel and adding a Power Query first[/]");
                    }
                }
                catch { }
                
                return 1;
            }

            // Display queries
            if (queries.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Query Name[/]");
                table.AddColumn("[bold]Formula (preview)[/]");

                foreach (var (name, formula) in queries.OrderBy(q => q.Name))
                {
                    table.AddRow(
                        $"[cyan]{name.EscapeMarkup()}[/]",
                        $"[dim]{(string.IsNullOrEmpty(formula) ? "(no formula)" : formula.EscapeMarkup())}[/]"
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine($"[bold]Total:[/] {queries.Count} Power Queries");
                
                // Provide usage hints for coding agents
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[dim]Next steps:[/]");
                AnsiConsole.MarkupLine($"[dim]• View query code:[/] [cyan]ExcelCLI pq-view \"{args[1]}\" \"QueryName\"[/]");
                AnsiConsole.MarkupLine($"[dim]• Export query:[/] [cyan]ExcelCLI pq-export \"{args[1]}\" \"QueryName\" \"output.pq\"[/]");
                AnsiConsole.MarkupLine($"[dim]• Refresh query:[/] [cyan]ExcelCLI pq-refresh \"{args[1]}\" \"QueryName\"[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No Power Queries found[/]");
                AnsiConsole.MarkupLine("[dim]Create one with:[/] [cyan]ExcelCLI pq-import \"{args[1]}\" \"QueryName\" \"code.pq\"[/]");
            }

            return 0;
        });
    }

    public int View(string[] args)
    {
        if (!ValidateArgs(args, 3, "pq-view <file.xlsx> <query-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            AnsiConsole.MarkupLine($"[yellow]Working Directory:[/] {Environment.CurrentDirectory}");
            AnsiConsole.MarkupLine($"[yellow]Full Path Expected:[/] {Path.GetFullPath(args[1])}");
            return 1;
        }

        var queryName = args[2];

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                // First, let's see what queries exist
                dynamic queriesCollection = workbook.Queries;
                int queryCount = queriesCollection.Count;
                
                AnsiConsole.MarkupLine($"[dim]Debug: Found {queryCount} queries in workbook[/]");
                
                dynamic? query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Query '{queryName.EscapeMarkup()}' not found");
                    
                    // Show available queries for coding agent context
                    if (queryCount > 0)
                    {
                        AnsiConsole.MarkupLine($"[yellow]Available queries in {Path.GetFileName(args[1])}:[/]");
                        
                        var availableQueries = new List<string>();
                        for (int i = 1; i <= queryCount; i++)
                        {
                            try
                            {
                                dynamic q = queriesCollection.Item(i);
                                string name = q.Name;
                                availableQueries.Add(name);
                                AnsiConsole.MarkupLine($"  [cyan]{i}.[/] {name.EscapeMarkup()}");
                            }
                            catch (Exception ex)
                            {
                                AnsiConsole.MarkupLine($"  [red]{i}.[/] <Error accessing query: {ex.Message.EscapeMarkup()}>");
                            }
                        }
                        
                        // Suggest closest match for coding agents
                        var closestMatch = FindClosestMatch(queryName, availableQueries);
                        if (!string.IsNullOrEmpty(closestMatch))
                        {
                            AnsiConsole.MarkupLine($"[yellow]Did you mean:[/] [cyan]{closestMatch}[/]");
                            AnsiConsole.MarkupLine($"[dim]Command suggestion:[/] [cyan]ExcelCLI pq-view \"{args[1]}\" \"{closestMatch}\"[/]");
                        }
                    }
                    else
                    {
                        AnsiConsole.MarkupLine("[yellow]No Power Queries found in this workbook[/]");
                        AnsiConsole.MarkupLine("[dim]Create one with:[/] [cyan]ExcelCLI pq-import file.xlsx \"QueryName\" \"code.pq\"[/]");
                    }
                    
                    return 1;
                }

                string formula = query.Formula;
                if (string.IsNullOrEmpty(formula))
                {
                    AnsiConsole.MarkupLine($"[yellow]Warning:[/] Query '{queryName.EscapeMarkup()}' has no formula content");
                    AnsiConsole.MarkupLine("[dim]This may be a function or connection-only query[/]");
                }

                AnsiConsole.MarkupLine($"[bold]Query:[/] [cyan]{queryName.EscapeMarkup()}[/]");
                AnsiConsole.MarkupLine($"[dim]Character count: {formula.Length:N0}[/]");
                AnsiConsole.WriteLine();
                
                var panel = new Panel(formula.EscapeMarkup())
                    .Header("[bold]Power Query M Code[/]")
                    .BorderColor(Color.Blue);
                    
                AnsiConsole.Write(panel);
                
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error accessing Power Query:[/] {ex.Message.EscapeMarkup()}");
                
                // Provide context for coding agents
                try
                {
                    dynamic queriesCollection = workbook.Queries;
                    AnsiConsole.MarkupLine($"[dim]Workbook has {queriesCollection.Count} total queries[/]");
                }
                catch
                {
                    AnsiConsole.MarkupLine("[dim]Unable to access Queries collection - workbook may not support Power Query[/]");
                }
                
                return 1;
            }
        });
    }

    public async Task<int> Update(string[] args)
    {
        if (!ValidateArgs(args, 4, "pq-update <file.xlsx> <query-name> <code.pq>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }
        if (!File.Exists(args[3]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Code file not found: {args[3]}");
            return 1;
        }

        var queryName = args[2];
        var newCode = await File.ReadAllTextAsync(args[3]);

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            dynamic? query = FindQuery(workbook, queryName);
            if (query == null)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Query '{queryName}' not found");
                return 1;
            }

            query.Formula = newCode;
            workbook.Save();
            AnsiConsole.MarkupLine($"[green]✓[/] Updated query '{queryName}'");
            return 0;
        });
    }

    public async Task<int> Export(string[] args)
    {
        if (!ValidateArgs(args, 4, "pq-export <file.xlsx> <query-name> <output.pq>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var queryName = args[2];
        var outputFile = args[3];

        return await Task.Run(() => WithExcel(args[1], false, async (excel, workbook) =>
        {
            dynamic? query = FindQuery(workbook, queryName);
            if (query == null)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Query '{queryName}' not found");
                return 1;
            }

            string formula = query.Formula;
            await File.WriteAllTextAsync(outputFile, formula);
            AnsiConsole.MarkupLine($"[green]✓[/] Exported query '{queryName}' to '{outputFile}'");
            return 0;
        }));
    }

    public async Task<int> Import(string[] args)
    {
        if (!ValidateArgs(args, 4, "pq-import <file.xlsx> <query-name> <source.pq>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }
        if (!File.Exists(args[3]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Source file not found: {args[3]}");
            return 1;
        }

        var queryName = args[2];
        var mCode = await File.ReadAllTextAsync(args[3]);

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            dynamic? existingQuery = FindQuery(workbook, queryName);

            if (existingQuery != null)
            {
                existingQuery.Formula = mCode;
                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Updated existing query '{queryName}'");
                return 0;
            }

            // Create new query
            dynamic queriesCollection = workbook.Queries;
            queriesCollection.Add(queryName, mCode, "");
            workbook.Save();
            AnsiConsole.MarkupLine($"[green]✓[/] Created new query '{queryName}'");
            return 0;
        });
    }

    public int Sources(string[] args)
    {
        if (!ValidateArgs(args, 2, "pq-sources <file.xlsx>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[bold]Excel.CurrentWorkbook() sources in:[/] {Path.GetFileName(args[1])}\n");
        AnsiConsole.MarkupLine("[dim]This shows what tables/ranges Power Query can see[/]\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            var sources = new List<(string Name, string Kind)>();

            // Create a temporary query to get Excel.CurrentWorkbook() results
            string diagnosticQuery = @"
let
    Sources = Excel.CurrentWorkbook()
in
    Sources";

            try
            {
                dynamic queriesCollection = workbook.Queries;

                // Create temp query
                dynamic tempQuery = queriesCollection.Add("_TempDiagnostic", diagnosticQuery, "");

                // Force refresh to evaluate
                tempQuery.Refresh();

                // Get the result (would need to read from cache/connection)
                // Since we can't easily get the result, let's parse from Excel tables instead

                // Clean up
                tempQuery.Delete();

                // Alternative: enumerate Excel objects directly
                // Get all tables from all worksheets
                dynamic worksheets = workbook.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    dynamic tables = worksheet.ListObjects;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        sources.Add((table.Name, "Table"));
                    }
                }

                // Get all named ranges
                dynamic names = workbook.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
                    string nameValue = name.Name;
                    if (!nameValue.StartsWith("_"))
                    {
                        sources.Add((nameValue, "Named Range"));
                    }
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message}");
                return 1;
            }

            // Display sources
            if (sources.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Name[/]");
                table.AddColumn("[bold]Kind[/]");

                foreach (var (name, kind) in sources.OrderBy(s => s.Name))
                {
                    table.AddRow(name, kind);
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Total: {sources.Count} sources[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No sources found[/]");
            }

            return 0;
        });
    }

    public int Test(string[] args)
    {
        if (!ValidateArgs(args, 3, "pq-test <file.xlsx> <source-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        string sourceName = args[2];
        AnsiConsole.MarkupLine($"[bold]Testing source:[/] {sourceName}\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                // Create a test query to load the source
                string testQuery = $@"
let
    Source = Excel.CurrentWorkbook(){{[Name=""{sourceName.Replace("\"", "\"\"")}""]]}}[Content]
in
    Source";

                dynamic queriesCollection = workbook.Queries;
                dynamic tempQuery = queriesCollection.Add("_TestQuery", testQuery, "");

                AnsiConsole.MarkupLine($"[green]✓[/] Source '[cyan]{sourceName}[/]' exists and can be loaded");
                AnsiConsole.MarkupLine($"\n[dim]Power Query M code to use:[/]");
                string mCode = $"Excel.CurrentWorkbook(){{{{[Name=\"{sourceName}\"]}}}}[Content]";
                var panel = new Panel(mCode.EscapeMarkup())
                {
                    Border = BoxBorder.Rounded,
                    BorderStyle = new Style(Color.Grey)
                };
                AnsiConsole.Write(panel);

                // Try to refresh
                try
                {
                    tempQuery.Refresh();
                    AnsiConsole.MarkupLine($"\n[green]✓[/] Query refreshes successfully");
                }
                catch
                {
                    AnsiConsole.MarkupLine($"\n[yellow]⚠[/] Could not refresh query (may need data source configuration)");
                }

                // Clean up
                tempQuery.Delete();

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]✗[/] Source '[cyan]{sourceName}[/]' not found or cannot be loaded");
                AnsiConsole.MarkupLine($"[dim]Error: {ex.Message}[/]\n");

                AnsiConsole.MarkupLine($"[yellow]Tip:[/] Use '[cyan]pq-sources[/]' to see all available sources");
                return 1;
            }
        });
    }

    public int Peek(string[] args)
    {
        if (!ValidateArgs(args, 3, "pq-peek <file.xlsx> <source-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        string sourceName = args[2];
        AnsiConsole.MarkupLine($"[bold]Preview of:[/] {sourceName}\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                // Check if it's a named range (single value)
                dynamic names = workbook.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
                    string nameValue = name.Name;
                    if (nameValue == sourceName)
                    {
                        try
                        {
                            var value = name.RefersToRange.Value;
                            AnsiConsole.MarkupLine($"[green]Named Range Value:[/] {value}");
                            AnsiConsole.MarkupLine($"[dim]Type: Single cell or range[/]");
                            return 0;
                        }
                        catch
                        {
                            AnsiConsole.MarkupLine($"[yellow]Named range found but value cannot be read (may be #REF!)[/]");
                            return 1;
                        }
                    }
                }

                // Check if it's a table
                dynamic worksheets = workbook.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    dynamic tables = worksheet.ListObjects;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        if (table.Name == sourceName)
                        {
                            int rowCount = table.ListRows.Count;
                            int colCount = table.ListColumns.Count;

                            AnsiConsole.MarkupLine($"[green]Table found:[/]");
                            AnsiConsole.MarkupLine($"  Rows: {rowCount}");
                            AnsiConsole.MarkupLine($"  Columns: {colCount}");

                            // Show column names
                            if (colCount > 0)
                            {
                                var columns = new List<string>();
                                dynamic listCols = table.ListColumns;
                                for (int c = 1; c <= Math.Min(colCount, 10); c++)
                                {
                                    columns.Add(listCols.Item(c).Name);
                                }
                                AnsiConsole.MarkupLine($"  Columns: {string.Join(", ", columns)}{(colCount > 10 ? "..." : "")}");
                            }

                            return 0;
                        }
                    }
                }

                AnsiConsole.MarkupLine($"[red]✗[/] Source '{sourceName}' not found");
                AnsiConsole.MarkupLine($"[yellow]Tip:[/] Use 'pq-sources' to see all available sources");
                return 1;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message}");
                return 1;
            }
        });
    }

    public int Eval(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-verify (file.xlsx) (m-expression)");
            Console.WriteLine("Example: pq-verify Plan.xlsx \"Excel.CurrentWorkbook(){[Name='Growth']}[Content]\"");
            AnsiConsole.MarkupLine("[dim]Purpose:[/] Validates Power Query M syntax and checks if expression can evaluate");
            return 1;
        }

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }
        string mExpression = args[2];
        AnsiConsole.MarkupLine($"[bold]Verifying Power Query M expression...[/]\n");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                // Create a temporary query with the expression
                string queryName = "_EvalTemp_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                dynamic queriesCollection = workbook.Queries;
                dynamic tempQuery = queriesCollection.Add(queryName, mExpression, "");

                // Try to refresh to evaluate
                try
                {
                    tempQuery.Refresh();

                    AnsiConsole.MarkupLine("[green]✓[/] Expression is valid and can evaluate\n");

                    // Try to get the result by creating a temporary worksheet and loading the query there
                    try
                    {
                        dynamic worksheets = workbook.Worksheets;
                        string tempSheetName = "_Eval_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                        dynamic tempSheet = worksheets.Add();
                        tempSheet.Name = tempSheetName;

                        // Use QueryTables.Add with WorkbookConnection
                        string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                        dynamic queryTables = tempSheet.QueryTables;

                        dynamic qt = queryTables.Add(
                            Connection: connString,
                            Destination: tempSheet.Range("A1")
                        );
                        qt.Refresh(BackgroundQuery: false);

                        // Read the value from A2 (A1 is header, A2 is data)
                        var resultValue = tempSheet.Range("A2").Value;

                        AnsiConsole.MarkupLine($"[dim]Expression:[/]");
                        var panel = new Panel(mExpression.EscapeMarkup())
                        {
                            Border = BoxBorder.Rounded,
                            BorderStyle = new Style(Color.Grey)
                        };
                        AnsiConsole.Write(panel);

                        string displayValue = resultValue != null ? resultValue.ToString() : "<null>";
                        AnsiConsole.MarkupLine($"\n[bold cyan]Result:[/] {displayValue.EscapeMarkup()}");

                        // Clean up
                        excel.DisplayAlerts = false;
                        tempSheet.Delete();
                        excel.DisplayAlerts = true;
                        tempQuery.Delete();
                        return 0;
                    }
                    catch
                    {
                        // If we can't load to sheet, just show that it evaluated
                        AnsiConsole.MarkupLine($"[dim]Expression:[/]");
                        var panel2 = new Panel(mExpression.EscapeMarkup())
                        {
                            Border = BoxBorder.Rounded,
                            BorderStyle = new Style(Color.Grey)
                        };
                        AnsiConsole.Write(panel2);

                        AnsiConsole.MarkupLine($"\n[green]✓[/] Syntax is valid and expression can evaluate");
                        AnsiConsole.MarkupLine($"[dim]Note:[/] Use 'sheet-read' to get actual values from Excel tables/ranges");
                        AnsiConsole.MarkupLine($"[dim]Tip:[/] Open Excel and check the query in Power Query Editor.");

                        // Clean up
                        tempQuery.Delete();
                        return 0;
                    }
                }
                catch (Exception evalEx)
                {
                    AnsiConsole.MarkupLine($"[red]✗[/] Expression evaluation failed");
                    AnsiConsole.MarkupLine($"[dim]Error: {evalEx.Message.EscapeMarkup()}[/]\n");

                    // Clean up
                    try { tempQuery.Delete(); } catch { }
                    return 1;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Refresh(string[] args)
    {
        if (!ValidateArgs(args, 2, "pq-refresh <file.xlsx> <query-name>"))
            return 1;

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Query name is required");
            AnsiConsole.MarkupLine("[dim]Usage: pq-refresh <file.xlsx> <query-name>[/]");
            return 1;
        }

        string queryName = args[2];

        AnsiConsole.MarkupLine($"[cyan]Refreshing query:[/] {queryName}");

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                // Find the query
                dynamic queriesCollection = workbook.Queries;
                dynamic? targetQuery = null;

                for (int i = 1; i <= queriesCollection.Count; i++)
                {
                    dynamic query = queriesCollection.Item(i);
                    if (query.Name == queryName)
                    {
                        targetQuery = query;
                        break;
                    }
                }

                if (targetQuery == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Query '{queryName}' not found");
                    return 1;
                }

                // Find the connection that uses this query and refresh it
                dynamic connections = workbook.Connections;
                bool refreshed = false;

                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic conn = connections.Item(i);

                    // Check if this connection is for our query
                    if (conn.Name.ToString().Contains(queryName))
                    {
                        AnsiConsole.MarkupLine($"[dim]Refreshing connection: {conn.Name}[/]");
                        conn.Refresh();
                        refreshed = true;
                        break;
                    }
                }

                if (!refreshed)
                {
                    // Check if this is a function (starts with "let" and defines a function parameter)
                    string formula = targetQuery.Formula;
                    bool isFunction = formula.Contains("(") && (formula.Contains("as table =>")
                                   || formula.Contains("as text =>")
                                   || formula.Contains("as number =>")
                                   || formula.Contains("as any =>"));

                    if (isFunction)
                    {
                        AnsiConsole.MarkupLine("[yellow]Note:[/] Query is a function - functions don't need refresh");
                        return 0;
                    }

                    // Try to refresh by finding connections that reference this query name
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic conn = connections.Item(i);
                        string connName = conn.Name.ToString();

                        // Connection names often match query names with underscores instead of spaces
                        string queryNameWithSpace = queryName.Replace("_", " ");

                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals(queryNameWithSpace, StringComparison.OrdinalIgnoreCase) ||
                            connName.Contains($"Query - {queryName}") ||
                            connName.Contains($"Query - {queryNameWithSpace}"))
                        {
                            AnsiConsole.MarkupLine($"[dim]Found connection: {connName}[/]");
                            conn.Refresh();
                            refreshed = true;
                            break;
                        }
                    }

                    if (!refreshed)
                    {
                        AnsiConsole.MarkupLine("[yellow]Note:[/] Query not loaded to a connection - may be an intermediate query");
                        AnsiConsole.MarkupLine("[dim]Try opening the file in Excel and refreshing manually[/]");
                    }
                }

                AnsiConsole.MarkupLine($"[green]√[/] Refreshed query '{queryName}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Errors(string[] args)
    {
        if (!ValidateArgs(args, 2, "pq-errors (file.xlsx) (query-name)"))
            return 1;

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        string? queryName = args.Length > 2 ? args[2] : null;

        AnsiConsole.MarkupLine(queryName != null
            ? $"[cyan]Checking errors for query:[/] {queryName}"
            : $"[cyan]Checking errors for all queries[/]");

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                dynamic queriesCollection = workbook.Queries;
                var errorsFound = new List<(string QueryName, string ErrorMessage)>();

                for (int i = 1; i <= queriesCollection.Count; i++)
                {
                    dynamic query = queriesCollection.Item(i);
                    string name = query.Name;

                    // Skip if filtering by specific query name
                    if (queryName != null && name != queryName)
                        continue;

                    try
                    {
                        // Try to access the formula - if there's a syntax error, this will throw
                        string formula = query.Formula;

                        // Check if the query has a connection with data
                        dynamic connections = workbook.Connections;
                        for (int j = 1; j <= connections.Count; j++)
                        {
                            dynamic conn = connections.Item(j);
                            if (conn.Name.ToString().Contains(name))
                            {
                                // Check for errors in the connection
                                try
                                {
                                    var oledbConnection = conn.OLEDBConnection;
                                    if (oledbConnection != null)
                                    {
                                        // Try to get background query state
                                        bool backgroundQuery = oledbConnection.BackgroundQuery;
                                    }
                                }
                                catch (Exception connEx)
                                {
                                    errorsFound.Add((name, connEx.Message));
                                }
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errorsFound.Add((name, ex.Message));
                    }
                }

                // Display errors
                if (errorsFound.Count > 0)
                {
                    AnsiConsole.MarkupLine($"\n[red]Found {errorsFound.Count} error(s):[/]\n");

                    var table = new Table();
                    table.AddColumn("[bold]Query Name[/]");
                    table.AddColumn("[bold]Error Message[/]");

                    foreach (var (name, error) in errorsFound)
                    {
                        table.AddRow(
                            name.EscapeMarkup(),
                            error.EscapeMarkup()
                        );
                    }

                    AnsiConsole.Write(table);
                    return 1;
                }
                else
                {
                    AnsiConsole.MarkupLine("[green]√[/] No errors found");
                    return 0;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int LoadTo(string[] args)
    {
        if (!ValidateArgs(args, 3, "pq-loadto <file.xlsx> <query-name> <sheet-name>"))
            return 1;

        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        string queryName = args[2];
        string sheetName = args[3];

        AnsiConsole.MarkupLine($"[cyan]Loading query '{queryName}' to sheet '{sheetName}'[/]");

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                // Find the query
                dynamic queriesCollection = workbook.Queries;
                dynamic? targetQuery = null;

                for (int i = 1; i <= queriesCollection.Count; i++)
                {
                    dynamic query = queriesCollection.Item(i);
                    if (query.Name == queryName)
                    {
                        targetQuery = query;
                        break;
                    }
                }

                if (targetQuery == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Query '{queryName}' not found");
                    return 1;
                }

                // Check if query is "Connection Only" by looking for existing connections or list objects that use it
                bool isConnectionOnly = true;
                string connectionName = "";

                // Check for existing connections
                dynamic connections = workbook.Connections;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic conn = connections.Item(i);
                    string connName = conn.Name.ToString();

                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        isConnectionOnly = false;
                        connectionName = connName;
                        break;
                    }
                }

                if (isConnectionOnly)
                {
                    AnsiConsole.MarkupLine($"[yellow]Note:[/] Query '{queryName}' is set to 'Connection Only'");
                    AnsiConsole.MarkupLine($"[dim]Will create table to load query data[/]");
                }
                else
                {
                    AnsiConsole.MarkupLine($"[dim]Query has existing connection: {connectionName}[/]");
                }

                // Check if sheet exists, if not create it
                dynamic sheets = workbook.Worksheets;
                dynamic? targetSheet = null;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    if (sheet.Name == sheetName)
                    {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    AnsiConsole.MarkupLine($"[dim]Creating new sheet: {sheetName}[/]");
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }
                else
                {
                    AnsiConsole.MarkupLine($"[dim]Using existing sheet: {sheetName}[/]");
                    // Clear existing content
                    targetSheet.Cells.Clear();
                }

                // Create a ListObject (Excel table) on the sheet
                AnsiConsole.MarkupLine($"[dim]Creating table from query[/]");

                try
                {
                    // Use QueryTables.Add method - the correct approach for Power Query
                    dynamic queryTables = targetSheet.QueryTables;

                    // The connection string for a Power Query uses Microsoft.Mashup.OleDb.1 provider
                    string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                    string commandText = $"SELECT * FROM [{queryName}]";

                    // Add the QueryTable
                    dynamic queryTable = queryTables.Add(
                        connectionString,
                        targetSheet.Range["A1"],
                        commandText
                    );

                    // Set properties
                    queryTable.Name = queryName.Replace(" ", "_");
                    queryTable.RefreshStyle = 1; // xlInsertDeleteCells

                    // Refresh the table to load data
                    AnsiConsole.MarkupLine($"[dim]Refreshing table data...[/]");
                    queryTable.Refresh(false);

                    AnsiConsole.MarkupLine($"[green]√[/] Query '{queryName}' loaded to sheet '{sheetName}'");
                    return 0;
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error creating table:[/] {ex.Message.EscapeMarkup()}");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    public int Delete(string[] args)
    {
        if (!ValidateArgs(args, 3, "pq-delete <file.xlsx> <query-name>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var queryName = args[2];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Query '{queryName}' not found");
                    return 1;
                }

                // Check if query is used by connections
                dynamic connections = workbook.Connections;
                var usingConnections = new List<string>();
                
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic conn = connections.Item(i);
                    string connName = conn.Name.ToString();
                    if (connName.Contains(queryName) || connName.Contains($"Query - {queryName}"))
                    {
                        usingConnections.Add(connName);
                    }
                }

                if (usingConnections.Count > 0)
                {
                    AnsiConsole.MarkupLine($"[yellow]Warning:[/] Query '{queryName}' is used by {usingConnections.Count} connection(s):");
                    foreach (var conn in usingConnections)
                    {
                        AnsiConsole.MarkupLine($"  - {conn.EscapeMarkup()}");
                    }
                    
                    var confirm = AnsiConsole.Confirm("Delete anyway? This may break dependent queries or worksheets.");
                    if (!confirm)
                    {
                        AnsiConsole.MarkupLine("[yellow]Cancelled[/]");
                        return 0;
                    }
                }

                // Delete the query
                query.Delete();
                workbook.Save();
                
                AnsiConsole.MarkupLine($"[green]✓[/] Deleted query '{queryName}'");
                
                if (usingConnections.Count > 0)
                {
                    AnsiConsole.MarkupLine("[yellow]Note:[/] You may need to refresh or recreate dependent connections");
                }
                
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
