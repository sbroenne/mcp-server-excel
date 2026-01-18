using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Table;

internal sealed class TableCommand : Command<TableCommand.Settings>
{
    private static readonly JsonSerializerOptions SortColumnsJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly ISessionService _sessionService;
    private readonly ITableCommands _tableCommands;
    private readonly ICliConsole _console;

    public TableCommand(ISessionService sessionService, ITableCommands tableCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _tableCommands = tableCommands ?? throw new ArgumentNullException(nameof(tableCommands));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            _console.WriteError("Session ID is required. Use 'session open' first.");
            return -1;
        }

        var action = settings.Action?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(action))
        {
            _console.WriteError("Action is required.");
            return -1;
        }

        var batch = _sessionService.GetBatch(settings.SessionId);

        return action switch
        {
            "list" => WriteResult(_tableCommands.List(batch)),
            "read" or "get" => ExecuteGet(batch, settings),
            "create" => ExecuteCreate(batch, settings),
            "rename" => ExecuteRename(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "resize" => ExecuteResize(batch, settings),
            "toggle-totals" => ExecuteToggleTotals(batch, settings),
            "set-column-total" => ExecuteSetColumnTotal(batch, settings),
            "append-rows" => ExecuteAppendRows(batch, settings),
            "get-data" => ExecuteGetData(batch, settings),
            "set-style" => ExecuteSetStyle(batch, settings),
            "add-to-data-model" => ExecuteAddToDataModel(batch, settings),
            "apply-filter" => ExecuteApplyFilter(batch, settings),
            "apply-filter-values" => ExecuteApplyFilterValues(batch, settings),
            "clear-filters" => ExecuteClearFilters(batch, settings),
            "get-filters" => ExecuteGetFilters(batch, settings),
            "add-column" => ExecuteAddColumn(batch, settings),
            "remove-column" => ExecuteRemoveColumn(batch, settings),
            "rename-column" => ExecuteRenameColumn(batch, settings),
            "get-structured-reference" => ExecuteGetStructuredReference(batch, settings),
            "sort" => ExecuteSort(batch, settings),
            "sort-multi" => ExecuteSortMulti(batch, settings),
            "get-column-number-format" or "get-column-format" => ExecuteGetColumnFormat(batch, settings),
            "set-column-number-format" or "set-column-format" => ExecuteSetColumnFormat(batch, settings),
            "create-from-dax" => ExecuteCreateFromDax(batch, settings),
            "update-dax" => ExecuteUpdateDax(batch, settings),
            "get-dax" => ExecuteGetDax(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteGet(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for get.");
            return -1;
        }

        return WriteResult(_tableCommands.Read(batch, settings.TableName));
    }

    private int ExecuteCreate(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) ||
            string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.Range))
        {
            _console.WriteError("--sheet, --table-name, and --range are required for create.");
            return -1;
        }

        var hasHeaders = settings.HasHeaders ?? true;
        try
        {
            _tableCommands.Create(batch, settings.SheetName, settings.TableName, settings.Range, hasHeaders, settings.TableStyle);
            _console.WriteInfo("Table created successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteRename(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.NewName))
        {
            _console.WriteError("--table-name and --new-name are required for rename.");
            return -1;
        }

        try
        {
            _tableCommands.Rename(batch, settings.TableName, settings.NewName);
            _console.WriteInfo("Table renamed successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for delete.");
            return -1;
        }

        try
        {
            _tableCommands.Delete(batch, settings.TableName);
            _console.WriteInfo("Table deleted successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteResize(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.Range))
        {
            _console.WriteError("--table-name and --range are required for resize.");
            return -1;
        }

        try
        {
            _tableCommands.Resize(batch, settings.TableName, settings.Range);
            _console.WriteInfo("Table resized successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteToggleTotals(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || settings.ShowTotals is null)
        {
            _console.WriteError("--table-name and --show-totals are required for toggle-totals.");
            return -1;
        }

        try
        {
            _tableCommands.ToggleTotals(batch, settings.TableName, settings.ShowTotals.Value);
            _console.WriteInfo("Table totals toggled successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetColumnTotal(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.ColumnName) ||
            string.IsNullOrWhiteSpace(settings.TotalFunction))
        {
            _console.WriteError("--table-name, --column-name, and --total-function are required for set-column-total.");
            return -1;
        }

        try
        {
            _tableCommands.SetColumnTotal(batch, settings.TableName, settings.ColumnName, settings.TotalFunction);
            _console.WriteInfo("Column total set successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteAppendRows(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for append-rows.");
            return -1;
        }

        var rows = LoadRows(settings);
        if (rows == null)
        {
            return -1;
        }

        try
        {
            _tableCommands.Append(batch, settings.TableName, rows);
            _console.WriteInfo("Rows appended successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetData(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for get-data.");
            return -1;
        }

        var visibleOnly = settings.VisibleOnly ?? false;
        return WriteResult(_tableCommands.GetData(batch, settings.TableName, visibleOnly));
    }

    private int ExecuteSetStyle(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.TableStyle))
        {
            _console.WriteError("--table-name and --table-style are required for set-style.");
            return -1;
        }

        try
        {
            _tableCommands.SetStyle(batch, settings.TableName, settings.TableStyle);
            _console.WriteInfo("Table style set successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteAddToDataModel(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for add-to-data-model.");
            return -1;
        }

        try
        {
            _tableCommands.AddToDataModel(batch, settings.TableName);
            _console.WriteInfo("Table added to data model successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteApplyFilter(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.ColumnName) ||
            string.IsNullOrWhiteSpace(settings.Criteria))
        {
            _console.WriteError("--table-name, --column-name, and --criteria are required for apply-filter.");
            return -1;
        }

        try
        {
            _tableCommands.ApplyFilter(batch, settings.TableName, settings.ColumnName, settings.Criteria);
            _console.WriteInfo("Filter applied successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteApplyFilterValues(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.ColumnName))
        {
            _console.WriteError("--table-name and --column-name are required for apply-filter-values.");
            return -1;
        }

        var values = SplitValues(settings.FilterValues);
        if (values == null || values.Count == 0)
        {
            _console.WriteError("Provide comma-separated values using --filter-values.");
            return -1;
        }

        try
        {
            _tableCommands.ApplyFilter(batch, settings.TableName, settings.ColumnName, values);
            _console.WriteInfo("Filter applied successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteClearFilters(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for clear-filters.");
            return -1;
        }

        try
        {
            _tableCommands.ClearFilters(batch, settings.TableName);
            _console.WriteInfo("Filters cleared successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetFilters(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for get-filters.");
            return -1;
        }

        return WriteResult(_tableCommands.GetFilters(batch, settings.TableName));
    }

    private int ExecuteAddColumn(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.ColumnName))
        {
            _console.WriteError("--table-name and --column-name are required for add-column.");
            return -1;
        }

        try
        {
            _tableCommands.AddColumn(batch, settings.TableName, settings.ColumnName, settings.ColumnPosition);
            _console.WriteInfo("Column added successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteRemoveColumn(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.ColumnName))
        {
            _console.WriteError("--table-name and --column-name are required for remove-column.");
            return -1;
        }

        try
        {
            _tableCommands.RemoveColumn(batch, settings.TableName, settings.ColumnName);
            _console.WriteInfo("Column removed successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteRenameColumn(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.ColumnName) ||
            string.IsNullOrWhiteSpace(settings.NewColumnName))
        {
            _console.WriteError("--table-name, --column-name, and --new-column-name are required for rename-column.");
            return -1;
        }

        try
        {
            _tableCommands.RenameColumn(batch, settings.TableName, settings.ColumnName, settings.NewColumnName);
            _console.WriteInfo("Column renamed successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetStructuredReference(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for get-structured-reference.");
            return -1;
        }

        if (!TryParseRegion(settings.Region, out var region))
        {
            _console.WriteError("--region is required and must be one of: all, data, headers, totals, thisrow.");
            return -1;
        }

        return WriteResult(_tableCommands.GetStructuredReference(batch, settings.TableName, region, settings.ColumnName));
    }

    private int ExecuteSort(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.ColumnName))
        {
            _console.WriteError("--table-name and --column-name are required for sort.");
            return -1;
        }

        var ascending = settings.SortAscending ?? true;
        try
        {
            _tableCommands.Sort(batch, settings.TableName, settings.ColumnName, ascending);
            _console.WriteInfo("Table sorted successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSortMulti(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for sort-multi.");
            return -1;
        }

        var sortColumns = LoadSortColumns(settings);
        if (sortColumns == null)
        {
            _console.WriteError("Provide sort columns using --sort-columns-json or --sort-columns-file.");
            return -1;
        }

        try
        {
            _tableCommands.Sort(batch, settings.TableName, sortColumns);
            _console.WriteInfo("Table sorted by multiple columns successfully");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetColumnFormat(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) || string.IsNullOrWhiteSpace(settings.ColumnName))
        {
            _console.WriteError("--table-name and --column-name are required for get-column-format.");
            return -1;
        }

        return WriteResult(_tableCommands.GetColumnNumberFormat(batch, settings.TableName, settings.ColumnName));
    }

    private int ExecuteSetColumnFormat(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.ColumnName) ||
            string.IsNullOrWhiteSpace(settings.FormatCode))
        {
            _console.WriteError("--table-name, --column-name, and --format-code are required for set-column-format.");
            return -1;
        }

        try
        {
            _tableCommands.SetColumnNumberFormat(batch, settings.TableName, settings.ColumnName, settings.FormatCode);
            _console.WriteInfo("Column format set successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteCreateFromDax(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) ||
            string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.DaxQuery))
        {
            _console.WriteError("--sheet, --table-name, and --dax-query are required for create-from-dax.");
            return -1;
        }

        var targetCell = settings.Range ?? "A1";
        try
        {
            _tableCommands.CreateFromDax(batch, settings.SheetName, settings.TableName, settings.DaxQuery, targetCell);
            _console.WriteInfo($"DAX-backed table '{settings.TableName}' created successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteUpdateDax(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.DaxQuery))
        {
            _console.WriteError("--table-name and --dax-query are required for update-dax.");
            return -1;
        }

        try
        {
            _tableCommands.UpdateDax(batch, settings.TableName, settings.DaxQuery);
            _console.WriteInfo($"DAX query for table '{settings.TableName}' updated successfully.");
            return 0;
        }
        catch (InvalidOperationException ex)
        {
            _console.WriteError($"Error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Unexpected error: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetDax(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName))
        {
            _console.WriteError("--table-name is required for get-dax.");
            return -1;
        }

        return WriteResult(_tableCommands.GetDax(batch, settings.TableName));
    }

    private List<List<object?>>? LoadRows(Settings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.RowsJson))
        {
            var parsed = ParseRowsJson(settings.RowsJson!);
            if (parsed == null)
            {
                _console.WriteError("Unable to parse --rows-json content.");
            }

            return parsed;
        }

        if (!string.IsNullOrWhiteSpace(settings.RowsFile))
        {
            if (!System.IO.File.Exists(settings.RowsFile))
            {
                _console.WriteError($"Rows file '{settings.RowsFile}' was not found.");
                return null;
            }

            var contents = System.IO.File.ReadAllText(settings.RowsFile);
            var parsed = ParseRowsJson(contents);
            if (parsed == null)
            {
                _console.WriteError($"Unable to parse JSON from '{settings.RowsFile}'.");
            }

            return parsed;
        }

        _console.WriteError("Provide table rows using --rows-json or --rows-file.");
        return null;
    }

    private static List<List<object?>>? ParseRowsJson(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<List<object?>>();
            foreach (var rowElement in document.RootElement.EnumerateArray())
            {
                if (rowElement.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                var row = new List<object?>();
                foreach (var cell in rowElement.EnumerateArray())
                {
                    row.Add(ConvertJsonCell(cell));
                }

                rows.Add(row);
            }

            return rows;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private List<TableSortColumn>? LoadSortColumns(Settings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.SortColumnsJson))
        {
            return ParseSortColumns(settings.SortColumnsJson!);
        }

        if (!string.IsNullOrWhiteSpace(settings.SortColumnsFile))
        {
            if (!System.IO.File.Exists(settings.SortColumnsFile))
            {
                _console.WriteError($"Sort file '{settings.SortColumnsFile}' was not found.");
                return null;
            }

            var contents = System.IO.File.ReadAllText(settings.SortColumnsFile);
            return ParseSortColumns(contents);
        }

        return null;
    }

    private static List<TableSortColumn>? ParseSortColumns(string json)
    {
        try
        {
            var columns = JsonSerializer.Deserialize<List<TableSortColumn>>(json, SortColumnsJsonOptions);

            return columns?.Count > 0 ? columns : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static object? ConvertJsonCell(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.Null => null,
            JsonValueKind.Number => element.TryGetInt64(out var i64) ? i64 : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String => element.GetString(),
            _ => element.GetRawText()
        };
    }

    private static bool TryParseRegion(string? input, out TableRegion region)
    {
        region = TableRegion.All;
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        var normalized = input.Replace("-", string.Empty, StringComparison.Ordinal).Trim();
        return Enum.TryParse(normalized, true, out region);
    }

    private static List<string>? SplitValues(string? values)
    {
        if (string.IsNullOrWhiteSpace(values))
        {
            return null;
        }

        var result = values
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();

        return result.Count > 0 ? result : null;
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown table action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <SHEET>")]
        public string? SheetName { get; init; }

        [CommandOption("--table-name <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewName { get; init; }

        [CommandOption("--range <RANGE>")]
        public string? Range { get; init; }

        [CommandOption("--has-headers <BOOL>")]
        public bool? HasHeaders { get; init; }

        [CommandOption("--table-style <STYLE>")]
        public string? TableStyle { get; init; }

        [CommandOption("--show-totals <BOOL>")]
        public bool? ShowTotals { get; init; }

        [CommandOption("--column-name <NAME>")]
        public string? ColumnName { get; init; }

        [CommandOption("--new-column-name <NAME>")]
        public string? NewColumnName { get; init; }

        [CommandOption("--column-position <NUMBER>")]
        public int? ColumnPosition { get; init; }

        [CommandOption("--total-function <FUNCTION>")]
        public string? TotalFunction { get; init; }

        [CommandOption("--criteria <CRITERIA>")]
        public string? Criteria { get; init; }

        [CommandOption("--filter-values <VALUES>")]
        public string? FilterValues { get; init; }

        [CommandOption("--rows-json <JSON>")]
        public string? RowsJson { get; init; }

        [CommandOption("--rows-file <PATH>")]
        public string? RowsFile { get; init; }

        [CommandOption("--region <REGION>")]
        public string? Region { get; init; }

        [CommandOption("--sort-ascending <BOOL>")]
        public bool? SortAscending { get; init; }

        [CommandOption("--sort-columns-json <JSON>")]
        public string? SortColumnsJson { get; init; }

        [CommandOption("--sort-columns-file <PATH>")]
        public string? SortColumnsFile { get; init; }

        [CommandOption("--format-code <CODE>")]
        public string? FormatCode { get; init; }

        [CommandOption("--visible-only <BOOL>")]
        public bool? VisibleOnly { get; init; }

        [CommandOption("--dax-query <QUERY>")]
        public string? DaxQuery { get; init; }
    }
}
