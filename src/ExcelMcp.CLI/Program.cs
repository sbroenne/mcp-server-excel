using System.Reflection;
using Sbroenne.ExcelMcp.CLI.Commands;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Set console encoding for better international character support
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        AnsiConsole.Write(new FigletText("Excel CLI").Color(Color.Blue));
        AnsiConsole.MarkupLine("[dim]Excel Command Line Interface for Coding Agents[/]\n");

        if (args.Length == 0)
        {
            ShowHelp();
            return 0;
        }

        // Input sanitization - prevent command injection
        if (args.Any(arg => string.IsNullOrWhiteSpace(arg)))
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Arguments cannot be empty or whitespace");
            return 1;
        }

        // Prevent excessively long arguments (potential DoS)
        const int MAX_ARG_LENGTH = 32767; // Windows path limit
        if (args.Any(arg => arg.Length > MAX_ARG_LENGTH))
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Argument too long (exceeds Windows path limit)");
            return 1;
        }

        try
        {
            var powerQuery = new PowerQueryCommands();
            var sheet = new SheetCommands();
            var range = new RangeCommands();
            var param = new ParameterCommands();
            var script = new ScriptCommands();
            var file = new FileCommands();
            var connection = new ConnectionCommands();
            var dataModel = new DataModelCommands();  // Used for dm-* commands
            var dataModelTom = new DataModelTomCommands();
            var table = new CliTableCommands();
            var pivot = new PivotTableCommands();

            return args[0].ToLower() switch
            {
                // Version and help commands
                "--version" or "-v" or "version" => ShowVersion(),

                // File commands
                "create-empty" => file.CreateEmpty(args),
                // Power Query commands
                "pq-list" => powerQuery.List(args),
                "pq-view" => powerQuery.View(args),
                "pq-update" => await powerQuery.Update(args),
                "pq-export" => await powerQuery.Export(args),
                "pq-import" => await powerQuery.Import(args),
                "pq-sources" => powerQuery.Sources(args),
                "pq-test" => powerQuery.Test(args),
                "pq-peek" => powerQuery.Peek(args),
                "pq-verify" => powerQuery.Eval(args),
                "pq-refresh" => powerQuery.Refresh(args),
                "pq-errors" => powerQuery.Errors(args),
                "pq-loadto" => powerQuery.LoadTo(args),
                "pq-delete" => powerQuery.Delete(args),

                // Power Query Load Configuration commands
                "pq-set-connection-only" => powerQuery.SetConnectionOnly(args),
                "pq-set-load-to-table" => powerQuery.SetLoadToTable(args),
                "pq-set-load-to-data-model" => powerQuery.SetLoadToDataModel(args),
                "pq-set-load-to-both" => powerQuery.SetLoadToBoth(args),
                "pq-get-load-config" => powerQuery.GetLoadConfig(args),

                // Sheet commands (lifecycle only - data operations moved to range-* commands in Phase 1A)
                "sheet-list" => sheet.List(args),
                "sheet-copy" => sheet.Copy(args),
                "sheet-delete" => sheet.Delete(args),
                "sheet-create" => sheet.Create(args),
                "sheet-rename" => sheet.Rename(args),

                // Range commands (data operations - replaces sheet-read/write/clear/append from Phase 1A)
                "range-get-values" => range.GetValues(args),
                "range-set-values" => range.SetValues(args),
                "range-get-formulas" => range.GetFormulas(args),
                "range-set-formulas" => range.SetFormulas(args),
                "range-clear-all" => range.ClearAll(args),
                "range-clear-contents" => range.ClearContents(args),
                "range-clear-formats" => range.ClearFormats(args),

                // Parameter commands
                "param-list" => param.List(args),
                "param-set" => param.Set(args),
                "param-get" => param.Get(args),
                "param-update" => param.Update(args),
                "param-create" => param.Create(args),
                "param-delete" => param.Delete(args),

                // Table commands
                "table-list" => table.List(args),
                "table-create" => table.Create(args),
                "table-rename" => table.Rename(args),
                "table-delete" => table.Delete(args),
                "table-info" => table.Info(args),
                "table-resize" => table.Resize(args),
                "table-toggle-totals" => table.ToggleTotals(args),
                "table-set-column-total" => table.SetColumnTotal(args),
                "table-append" => table.AppendRows(args),
                "table-set-style" => table.SetStyle(args),
                "table-add-to-datamodel" => table.AddToDataModel(args),
                "table-apply-filter" => table.ApplyFilter(args),
                "table-apply-filter-values" => table.ApplyFilterValues(args),
                "table-clear-filters" => table.ClearFilters(args),
                "table-get-filters" => table.GetFilters(args),
                "table-add-column" => table.AddColumn(args),
                "table-remove-column" => table.RemoveColumn(args),
                "table-rename-column" => table.RenameColumn(args),
                "table-get-structured-reference" => table.GetStructuredReference(args),
                "table-sort" => table.Sort(args),
                "table-sort-multi" => table.SortMulti(args),

                // PivotTable commands
                "pivot-list" => pivot.List(args),
                "pivot-create-from-range" => pivot.CreateFromRange(args),
                "pivot-add-row-field" => pivot.AddRowField(args),
                "pivot-add-value-field" => pivot.AddValueField(args),
                "pivot-refresh" => pivot.Refresh(args),

                // Connection commands
                "conn-list" => connection.List(args),
                "conn-view" => connection.View(args),
                "conn-import" => connection.Import(args),
                "conn-export" => connection.Export(args),
                "conn-update" => connection.Update(args),
                "conn-refresh" => connection.Refresh(args),
                "conn-delete" => connection.Delete(args),
                "conn-loadto" => connection.LoadTo(args),
                "conn-properties" => connection.GetProperties(args),
                "conn-set-properties" => connection.SetProperties(args),
                "conn-test" => connection.Test(args),

                // Script commands
                "script-list" => script.List(args),
                "script-view" => script.View(args),
                "script-export" => script.Export(args),
                "script-import" => await script.Import(args),
                "script-update" => await script.Update(args),
                "script-run" => script.Run(args),

                // Data Model commands (READ operations via COM API)
                "dm-list-tables" => dataModel.ListTables(args),
                "dm-list-measures" => dataModel.ListMeasures(args),
                "dm-view-measure" => dataModel.ViewMeasure(args),
                "dm-export-measure" => dataModel.ExportMeasure(args),
                "dm-list-relationships" => dataModel.ListRelationships(args),
                "dm-refresh" => dataModel.Refresh(args),
                "dm-delete-measure" => dataModel.DeleteMeasure(args),
                "dm-delete-relationship" => dataModel.DeleteRelationship(args),

                // Data Model Phase 2 commands (Discovery operations via COM API)
                "dm-list-columns" => dataModel.ListColumns(args),
                "dm-view-table" => dataModel.ViewTable(args),
                "dm-get-model-info" => dataModel.GetModelInfo(args),

                // Data Model Phase 2 commands (CREATE/UPDATE operations via COM API)
                "dm-create-measure" => dataModel.CreateMeasure(args),
                "dm-update-measure" => dataModel.UpdateMeasure(args),
                "dm-create-relationship" => dataModel.CreateRelationship(args),
                "dm-update-relationship" => dataModel.UpdateRelationship(args),

                // Data Model TOM (Tabular Object Model) commands - Advanced CRUD operations (future)
                "dm-create-column" => dataModelTom.CreateCalculatedColumn(args),
                "dm-view-column" => dataModelTom.ViewCalculatedColumn(args),
                "dm-update-column" => dataModelTom.UpdateCalculatedColumn(args),
                "dm-delete-column" => dataModelTom.DeleteCalculatedColumn(args),
                "dm-validate-dax" => dataModelTom.ValidateDax(args),

                "--help" or "-h" => ShowHelp(),
                _ => ShowHelp()
            };
        }
        catch (Exception ex)
        {
            // Enhanced error reporting for coding agents
            AnsiConsole.MarkupLine($"[red]Fatal Error:[/] {ex.Message.EscapeMarkup()}");

            // Provide specific guidance based on error type
            if (ex is FileNotFoundException fnfEx)
            {
                AnsiConsole.MarkupLine($"[yellow]File not found:[/] {fnfEx.FileName.EscapeMarkup()}");
                AnsiConsole.MarkupLine($"[yellow]Working Directory:[/] {Environment.CurrentDirectory}");
                if (!string.IsNullOrEmpty(fnfEx.FileName))
                {
                    AnsiConsole.MarkupLine($"[yellow]Expected Path:[/] {Path.GetFullPath(fnfEx.FileName)}");
                }
            }
            else if (ex is UnauthorizedAccessException)
            {
                AnsiConsole.MarkupLine("[yellow]Access denied. Try:[/]");
                AnsiConsole.MarkupLine("  • Run as Administrator");
                AnsiConsole.MarkupLine("  • Check file permissions");
                AnsiConsole.MarkupLine("  • Close Excel if file is open");
            }
            else if (ex is InvalidOperationException && ex.Message.Contains("Excel"))
            {
                AnsiConsole.MarkupLine("[yellow]Excel issue detected. Try:[/]");
                AnsiConsole.MarkupLine("  • Verify Excel is installed");
                AnsiConsole.MarkupLine("  • Close all Excel instances");
                AnsiConsole.MarkupLine("  • Run Excel repair from Control Panel");
                AnsiConsole.MarkupLine("  • Check Windows Updates");
            }

            // Show command context if available
            if (args.Length > 0)
            {
                AnsiConsole.MarkupLine($"[dim]Command attempted:[/] [cyan]{string.Join(" ", args.Select(a => a.Contains(' ') ? $"\"{a}\"" : a))}[/]");
            }

            // In debug builds or if verbose flag, show full details
            bool showDetails = args.Contains("--verbose") || args.Contains("-v") ||
                              Environment.GetEnvironmentVariable("EXCELCLI_DEBUG") == "1";

            if (showDetails)
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[dim]=== DETAILED ERROR INFORMATION ===[/]");
                AnsiConsole.MarkupLine($"[dim]Exception Type:[/] {ex.GetType().FullName}");
                AnsiConsole.MarkupLine($"[dim]HResult:[/] 0x{ex.HResult:X8}");

                if (ex.Data.Count > 0)
                {
                    AnsiConsole.MarkupLine("[dim]Exception Data:[/]");
                    foreach (var key in ex.Data.Keys)
                    {
                        AnsiConsole.MarkupLine($"[dim]  {key}:[/] {ex.Data[key]}");
                    }
                }

                if (ex.InnerException != null)
                {
                    AnsiConsole.MarkupLine($"[dim]Inner Exception:[/] {ex.InnerException.GetType().Name}");
                    AnsiConsole.MarkupLine($"[dim]Inner Message:[/] {ex.InnerException.Message.EscapeMarkup()}");
                }

                AnsiConsole.MarkupLine("[dim]Stack Trace:[/]");
                AnsiConsole.MarkupLine($"[dim]{ex.StackTrace?.EscapeMarkup()}[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[dim]For detailed error information, add [cyan]--verbose[/] flag[/]");
            }

            return 1;
        }
    }

    static int ShowVersion()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var informationalVersion = System.Reflection.Assembly.GetExecutingAssembly()
            .GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? version?.ToString() ?? "Unknown";

        AnsiConsole.MarkupLine($"[bold cyan]ExcelMcp.CLI[/] [green]v{informationalVersion}[/]");
        AnsiConsole.MarkupLine("[dim]Excel Command Line Interface for Coding Agents[/]");
        AnsiConsole.MarkupLine($"[dim]Runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}[/]");
        AnsiConsole.MarkupLine($"[dim]Platform: {System.Runtime.InteropServices.RuntimeInformation.OSDescription}[/]");
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[bold]Repository:[/] https://github.com/sbroenne/mcp-server-excel");
        AnsiConsole.MarkupLine("[bold]License:[/] MIT");

        return 0;
    }

    static int ShowHelp()
    {
        AnsiConsole.Write(new Rule("[bold cyan]ExcelMcp.CLI - Excel Command Line Interface for Coding Agents[/]").RuleStyle("grey"));
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold]Usage:[/] excelcli command args");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]File Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]create-empty[/] file.xlsx                      Create empty Excel workbook");
        AnsiConsole.MarkupLine("  [cyan]create-empty[/] file.xlsm                      Create macro-enabled workbook");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Power Query Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]pq-list[/] file.xlsx                           List all Power Queries");
        AnsiConsole.MarkupLine("  [cyan]pq-view[/] file.xlsx query-name               View Power Query M code");
        AnsiConsole.MarkupLine("  [cyan]pq-update[/] file.xlsx query-name code.pq     Update Power Query from file");
        AnsiConsole.MarkupLine("    Options: [dim]--privacy-level <None|Private|Organizational|Public>[/]");
        AnsiConsole.MarkupLine("  [cyan]pq-export[/] file.xlsx query-name out.pq      Export Power Query to file");
        AnsiConsole.MarkupLine("  [cyan]pq-import[/] file.xlsx query-name src.pq      Import/create Power Query");
        AnsiConsole.MarkupLine("    Options: [dim]--privacy-level <None|Private|Organizational|Public> --connection-only[/]");
        AnsiConsole.MarkupLine("  [cyan]pq-refresh[/] file.xlsx query-name            Refresh a specific Power Query");
        AnsiConsole.MarkupLine("  [cyan]pq-loadto[/] file.xlsx query-name sheet       Load Power Query to worksheet");
        AnsiConsole.MarkupLine("  [cyan]pq-delete[/] file.xlsx query-name             Delete Power Query");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Power Query Load Configuration:[/]");
        AnsiConsole.MarkupLine("  [cyan]pq-set-connection-only[/] file.xlsx query     Set query to Connection Only");
        AnsiConsole.MarkupLine("  [cyan]pq-set-load-to-table[/] file.xlsx query sheet Set query to Load to Table");
        AnsiConsole.MarkupLine("  [cyan]pq-set-load-to-data-model[/] file.xlsx query Set query to Load to Data Model");
        AnsiConsole.MarkupLine("  [cyan]pq-set-load-to-both[/] file.xlsx query sheet Set query to Load to Both");
        AnsiConsole.MarkupLine("  [cyan]pq-get-load-config[/] file.xlsx query        Get current load configuration");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Sheet Commands (Lifecycle Management):[/]");
        AnsiConsole.MarkupLine("  [cyan]sheet-list[/] file.xlsx                         List all worksheets");
        AnsiConsole.MarkupLine("  [cyan]sheet-copy[/] file.xlsx src-sheet new-sheet     Copy worksheet");
        AnsiConsole.MarkupLine("  [cyan]sheet-delete[/] file.xlsx sheet-name            Delete worksheet");
        AnsiConsole.MarkupLine("  [cyan]sheet-create[/] file.xlsx sheet-name            Create new worksheet");
        AnsiConsole.MarkupLine("  [cyan]sheet-rename[/] file.xlsx old-name new-name     Rename worksheet");
        AnsiConsole.MarkupLine("  [dim]Note: Data operations (read, write, clear) moved to range-* commands[/]");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Range Commands (Data Operations):[/]");
        AnsiConsole.MarkupLine("  [cyan]range-get-values[/] file.xlsx sheet range      Read values from range (output: CSV)");
        AnsiConsole.MarkupLine("  [cyan]range-set-values[/] file.xlsx sheet range csv  Write CSV data to range");
        AnsiConsole.MarkupLine("  [cyan]range-get-formulas[/] file.xlsx sheet range    Read formulas from range");
        AnsiConsole.MarkupLine("  [cyan]range-set-formulas[/] file.xlsx sheet range csv Set formulas from CSV");
        AnsiConsole.MarkupLine("  [cyan]range-clear-all[/] file.xlsx sheet range       Clear all (values, formulas, formats)");
        AnsiConsole.MarkupLine("  [cyan]range-clear-contents[/] file.xlsx sheet range  Clear contents (preserve formats)");
        AnsiConsole.MarkupLine("  [cyan]range-clear-formats[/] file.xlsx sheet range   Clear formats (preserve values)");
        AnsiConsole.MarkupLine("  [dim]Note: Single cell = 1x1 range (e.g., A1). Named ranges: use empty sheet \"\"[/]");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Parameter Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]param-list[/] file.xlsx                        List all named ranges");
        AnsiConsole.MarkupLine("  [cyan]param-get[/] file.xlsx param-name             Get named range value");
        AnsiConsole.MarkupLine("  [cyan]param-set[/] file.xlsx param-name value        Set named range value");
        AnsiConsole.MarkupLine("  [cyan]param-update[/] file.xlsx param-name ref       Update named range reference");
        AnsiConsole.MarkupLine("  [cyan]param-create[/] file.xlsx param-name ref       Create named range");
        AnsiConsole.MarkupLine("  [cyan]param-delete[/] file.xlsx param-name           Delete named range");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Table Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]table-list[/] file.xlsx                        List all tables");
        AnsiConsole.MarkupLine("  [cyan]table-create[/] file.xlsx sheet name range    Create table from range");
        AnsiConsole.MarkupLine("  [cyan]table-info[/] file.xlsx table-name            Get table details");
        AnsiConsole.MarkupLine("  [cyan]table-rename[/] file.xlsx old-name new-name   Rename table");
        AnsiConsole.MarkupLine("  [cyan]table-delete[/] file.xlsx table-name          Delete table");
        AnsiConsole.MarkupLine("  [cyan]table-resize[/] file.xlsx table-name range    Resize table");
        AnsiConsole.MarkupLine("  [cyan]table-set-style[/] file.xlsx table-name style Change table style");
        AnsiConsole.MarkupLine("  [cyan]table-toggle-totals[/] file.xlsx table-name true|false  Show/hide totals");
        AnsiConsole.MarkupLine("  [cyan]table-set-column-total[/] file.xlsx table col func  Set column total function");
        AnsiConsole.MarkupLine("  [cyan]table-append[/] file.xlsx table-name data.csv Append rows to table");
        AnsiConsole.MarkupLine("  [cyan]table-add-to-datamodel[/] file.xlsx table-name  Add table to Data Model");
        AnsiConsole.MarkupLine("  [cyan]table-apply-filter[/] file.xlsx table col criteria  Filter by criteria");
        AnsiConsole.MarkupLine("  [cyan]table-apply-filter-values[/] file.xlsx table col vals  Filter by values");
        AnsiConsole.MarkupLine("  [cyan]table-clear-filters[/] file.xlsx table-name   Clear all filters");
        AnsiConsole.MarkupLine("  [cyan]table-get-filters[/] file.xlsx table-name    Get filter state");
        AnsiConsole.MarkupLine("  [cyan]table-add-column[/] file.xlsx table col [[pos]]  Add column");
        AnsiConsole.MarkupLine("  [cyan]table-remove-column[/] file.xlsx table col    Remove column");
        AnsiConsole.MarkupLine("  [cyan]table-rename-column[/] file.xlsx table old new  Rename column");
        AnsiConsole.MarkupLine("  [cyan]table-get-structured-reference[/] file.xlsx table region [[col]]  Get ref");
        AnsiConsole.MarkupLine("  [cyan]table-sort[/] file.xlsx table col [[asc|desc]]  Sort by column");
        AnsiConsole.MarkupLine("  [cyan]table-sort-multi[/] file.xlsx table col1:asc col2:desc...  Multi-sort");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]PivotTable Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]pivot-list[/] file.xlsx                          List all PivotTables");
        AnsiConsole.MarkupLine("  [cyan]pivot-create-from-range[/] file.xlsx src-sheet src-range dest-sheet dest-cell name");
        AnsiConsole.MarkupLine("    [dim]Example: pivot-create-from-range sales.xlsx Data A1:D100 Analysis A1 SalesPivot[/]");
        AnsiConsole.MarkupLine("  [cyan]pivot-add-row-field[/] file.xlsx pivot-name field [[position]]");
        AnsiConsole.MarkupLine("  [cyan]pivot-add-value-field[/] file.xlsx pivot-name field [[function]] [[custom-name]]");
        AnsiConsole.MarkupLine("    [dim]Functions: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, VarP[/]");
        AnsiConsole.MarkupLine("  [cyan]pivot-refresh[/] file.xlsx pivot-name           Refresh PivotTable data");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Connection Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]conn-list[/] file.xlsx                         List all connections");
        AnsiConsole.MarkupLine("  [cyan]conn-view[/] file.xlsx conn-name              View connection details");
        AnsiConsole.MarkupLine("  [cyan]conn-import[/] file.xlsx conn-name def.json   Import connection from JSON");
        AnsiConsole.MarkupLine("  [cyan]conn-export[/] file.xlsx conn-name out.json   Export connection to JSON");
        AnsiConsole.MarkupLine("  [cyan]conn-update[/] file.xlsx conn-name def.json   Update connection from JSON");
        AnsiConsole.MarkupLine("  [cyan]conn-refresh[/] file.xlsx conn-name           Refresh connection data");
        AnsiConsole.MarkupLine("  [cyan]conn-delete[/] file.xlsx conn-name            Delete connection");
        AnsiConsole.MarkupLine("  [cyan]conn-loadto[/] file.xlsx conn-name sheet      Load connection to worksheet");
        AnsiConsole.MarkupLine("  [cyan]conn-properties[/] file.xlsx conn-name        Get connection properties");
        AnsiConsole.MarkupLine("  [cyan]conn-set-properties[/] file.xlsx conn-name... Set connection properties");
        AnsiConsole.MarkupLine("  [cyan]conn-test[/] file.xlsx conn-name              Test connection validity");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Script Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]script-list[/] file.xlsm                       List all VBA scripts");
        AnsiConsole.MarkupLine("  [cyan]script-view[/] file.xlsm module-name           View VBA module code");
        AnsiConsole.MarkupLine("  [cyan]script-export[/] file.xlsm script (file)       Export VBA script");
        AnsiConsole.MarkupLine("  [cyan]script-import[/] file.xlsm module-name vba.txt Import VBA script");
        AnsiConsole.MarkupLine("  [cyan]script-update[/] file.xlsm module-name vba.txt Update VBA script");
        AnsiConsole.MarkupLine("  [cyan]script-run[/] file.xlsm macro-name (params)    Run VBA macro");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Data Model Commands:[/]");
        AnsiConsole.MarkupLine("  [cyan]dm-list-tables[/] file.xlsx                    List all Data Model tables");
        AnsiConsole.MarkupLine("  [cyan]dm-list-measures[/] file.xlsx                  List all DAX measures");
        AnsiConsole.MarkupLine("  [cyan]dm-view-measure[/] file.xlsx measure-name     View DAX measure formula");
        AnsiConsole.MarkupLine("  [cyan]dm-export-measure[/] file.xlsx measure out.dax Export DAX measure to file");
        AnsiConsole.MarkupLine("  [cyan]dm-list-relationships[/] file.xlsx            List Data Model relationships");
        AnsiConsole.MarkupLine("  [cyan]dm-refresh[/] file.xlsx                        Refresh Data Model");
        AnsiConsole.MarkupLine("  [cyan]dm-delete-measure[/] file.xlsx measure-name   Delete DAX measure");
        AnsiConsole.MarkupLine("  [cyan]dm-delete-relationship[/] file.xlsx from-tbl from-col to-tbl to-col  Delete relationship");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold yellow]Data Model TOM Commands (Advanced CRUD):[/]");
        AnsiConsole.MarkupLine("  [cyan]dm-create-measure[/] file.xlsx table name formula  Create DAX measure");
        AnsiConsole.MarkupLine("  [cyan]dm-update-measure[/] file.xlsx name [[options]]      Update DAX measure");
        AnsiConsole.MarkupLine("  [cyan]dm-create-relationship[/] file.xlsx from to        Create table relationship");
        AnsiConsole.MarkupLine("  [cyan]dm-update-relationship[/] file.xlsx from to [[opts]] Update relationship");
        AnsiConsole.MarkupLine("  [cyan]dm-create-column[/] file.xlsx table name formula   Create calculated column");
        AnsiConsole.MarkupLine("  [cyan]dm-list-columns[/] file.xlsx [[table]]               List calculated columns");
        AnsiConsole.MarkupLine("  [cyan]dm-view-column[/] file.xlsx table column           View column details");
        AnsiConsole.MarkupLine("  [cyan]dm-update-column[/] file.xlsx table column [[opts]]  Update calculated column");
        AnsiConsole.MarkupLine("  [cyan]dm-delete-column[/] file.xlsx table column         Delete calculated column");
        AnsiConsole.MarkupLine("  [cyan]dm-validate-dax[/] file.xlsx formula               Validate DAX syntax");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold green]Examples:[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli create-empty \"Plan.xlsm\"[/]            [dim]# Create macro-enabled workbook[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli script-import \"Plan.xlsm\" \"Helper\" \"code.vba\"[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli pq-list \"Plan.xlsx\"[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli pq-view \"Plan.xlsx\" \"Milestones\"[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli pq-import \"Plan.xlsx\" \"fnHelper\" \"function.pq\"[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli sheet-list \"Plan.xlsx\"[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli range-get-values \"Plan.xlsx\" \"Data\" \"A1:D10\"[/]");
        AnsiConsole.MarkupLine("  [dim]excelcli param-set \"Plan.xlsx\" \"Start_Date\" \"2025-01-01\"[/]");
        AnsiConsole.WriteLine();

        AnsiConsole.MarkupLine("[bold]Requirements:[/] Windows + Excel + .NET 10.0");
        AnsiConsole.MarkupLine("[bold]License:[/] MIT | [bold]Repository:[/] https://github.com/sbroenne/mcp-server-excel");

        return 0;
    }
}
