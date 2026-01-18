using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.PivotTable;

internal sealed class PivotTableCommand : Command<PivotTableCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IPivotTableCommands _pivotTableCommands;
    private readonly ICliConsole _console;

    public PivotTableCommand(ISessionService sessionService, IPivotTableCommands pivotTableCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _pivotTableCommands = pivotTableCommands ?? throw new ArgumentNullException(nameof(pivotTableCommands));
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
            "list" => WriteResult(_pivotTableCommands.List(batch)),
            "get" => ExecuteGet(batch, settings),
            "create-from-range" => ExecuteCreateFromRange(batch, settings),
            "create-from-table" => ExecuteCreateFromTable(batch, settings),
            "create-from-datamodel" => ExecuteCreateFromDataModel(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "refresh" => ExecuteRefresh(batch, settings),
            "list-fields" => ExecuteListFields(batch, settings),
            "add-row-field" => ExecuteAddRowField(batch, settings),
            "add-column-field" => ExecuteAddColumnField(batch, settings),
            "add-value-field" => ExecuteAddValueField(batch, settings),
            "add-filter-field" => ExecuteAddFilterField(batch, settings),
            "remove-field" => ExecuteRemoveField(batch, settings),
            "set-field-function" => ExecuteSetFieldFunction(batch, settings),
            "set-field-name" => ExecuteSetFieldName(batch, settings),
            "set-field-format" => ExecuteSetFieldFormat(batch, settings),
            "set-field-filter" => ExecuteSetFieldFilter(batch, settings),
            "sort-field" => ExecuteSortField(batch, settings),
            "group-by-date" => ExecuteGroupByDate(batch, settings),
            "group-by-numeric" => ExecuteGroupByNumeric(batch, settings),
            "create-calculated-field" => ExecuteCreateCalculatedField(batch, settings),
            "list-calculated-fields" => ExecuteListCalculatedFields(batch, settings),
            "delete-calculated-field" => ExecuteDeleteCalculatedField(batch, settings),
            "list-calculated-members" => ExecuteListCalculatedMembers(batch, settings),
            "create-calculated-member" => ExecuteCreateCalculatedMember(batch, settings),
            "delete-calculated-member" => ExecuteDeleteCalculatedMember(batch, settings),
            "set-layout" => ExecuteSetLayout(batch, settings),
            "set-subtotals" => ExecuteSetSubtotals(batch, settings),
            "set-grand-totals" => ExecuteSetGrandTotals(batch, settings),
            "get-data" => ExecuteGetData(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteGet(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for get.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.Read(batch, settings.PivotTableName));
    }

    private int ExecuteCreateFromRange(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) ||
            string.IsNullOrWhiteSpace(settings.Range) ||
            string.IsNullOrWhiteSpace(settings.DestinationSheet) ||
            string.IsNullOrWhiteSpace(settings.DestinationCell) ||
            string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--sheet, --range, --destination-sheet, --destination-cell, and --pivot-name are required for create-from-range.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.CreateFromRange(
            batch,
            settings.SheetName,
            settings.Range,
            settings.DestinationSheet,
            settings.DestinationCell,
            settings.PivotTableName));
    }

    private int ExecuteCreateFromTable(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TableName) ||
            string.IsNullOrWhiteSpace(settings.DestinationSheet) ||
            string.IsNullOrWhiteSpace(settings.DestinationCell) ||
            string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--table-name, --destination-sheet, --destination-cell, and --pivot-name are required for create-from-table.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.CreateFromTable(
            batch,
            settings.TableName,
            settings.DestinationSheet,
            settings.DestinationCell,
            settings.PivotTableName));
    }

    private int ExecuteCreateFromDataModel(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.DataModelTableName) ||
            string.IsNullOrWhiteSpace(settings.DestinationSheet) ||
            string.IsNullOrWhiteSpace(settings.DestinationCell) ||
            string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--data-model-table, --destination-sheet, --destination-cell, and --pivot-name are required for create-from-datamodel.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.CreateFromDataModel(
            batch,
            settings.DataModelTableName,
            settings.DestinationSheet,
            settings.DestinationCell,
            settings.PivotTableName));
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for delete.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.Delete(batch, settings.PivotTableName));
    }

    private int ExecuteRefresh(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for refresh.");
            return -1;
        }

        if (!TryGetTimeout(settings.TimeoutSeconds, out var timeout))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.Refresh(batch, settings.PivotTableName, timeout));
    }

    private int ExecuteListFields(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for list-fields.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.ListFields(batch, settings.PivotTableName));
    }

    private int ExecuteAddRowField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "add-row-field"))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.AddRowField(batch, settings.PivotTableName!, settings.FieldName!, settings.Position));
    }

    private int ExecuteAddColumnField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "add-column-field"))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.AddColumnField(batch, settings.PivotTableName!, settings.FieldName!, settings.Position));
    }

    private int ExecuteAddValueField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "add-value-field"))
        {
            return -1;
        }

        if (!TryParseAggregation(settings.Aggregation, out var aggregationFunction))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.AddValueField(batch, settings.PivotTableName!, settings.FieldName!, aggregationFunction, settings.CustomName));
    }

    private int ExecuteAddFilterField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "add-filter-field"))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.AddFilterField(batch, settings.PivotTableName!, settings.FieldName!));
    }

    private int ExecuteRemoveField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "remove-field"))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.RemoveField(batch, settings.PivotTableName!, settings.FieldName!));
    }

    private int ExecuteSetFieldFunction(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "set-field-function"))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.Aggregation))
        {
            _console.WriteError("--aggregation is required for set-field-function.");
            return -1;
        }

        if (!TryParseAggregation(settings.Aggregation, out var aggregationFunction))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.SetFieldFunction(batch, settings.PivotTableName!, settings.FieldName!, aggregationFunction));
    }

    private int ExecuteSetFieldName(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "set-field-name"))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.CustomName))
        {
            _console.WriteError("--custom-name is required for set-field-name.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.SetFieldName(batch, settings.PivotTableName!, settings.FieldName!, settings.CustomName));
    }

    private int ExecuteSetFieldFormat(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "set-field-format"))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.NumberFormat))
        {
            _console.WriteError("--number-format is required for set-field-format.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.SetFieldFormat(batch, settings.PivotTableName!, settings.FieldName!, settings.NumberFormat));
    }

    private int ExecuteSetFieldFilter(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "set-field-filter"))
        {
            return -1;
        }

        var filterValues = LoadFilterValues(settings);
        if (filterValues == null)
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.SetFieldFilter(batch, settings.PivotTableName!, settings.FieldName!, filterValues));
    }

    private int ExecuteSortField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "sort-field"))
        {
            return -1;
        }

        if (!TryParseSortDirection(settings.SortDirection, out var direction))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.SortField(batch, settings.PivotTableName!, settings.FieldName!, direction));
    }

    private int ExecuteGetData(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for get-data.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.GetData(batch, settings.PivotTableName));
    }

    private int ExecuteGroupByDate(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "group-by-date"))
        {
            return -1;
        }

        if (!TryParseDateGroupingInterval(settings.DateGroupingInterval, out var interval))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.GroupByDate(batch, settings.PivotTableName!, settings.FieldName!, interval));
    }

    private int ExecuteGroupByNumeric(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "group-by-numeric"))
        {
            return -1;
        }

        if (!settings.NumericGroupingInterval.HasValue || settings.NumericGroupingInterval.Value <= 0)
        {
            _console.WriteError("--numeric-interval is required and must be > 0 for group-by-numeric.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.GroupByNumeric(
            batch,
            settings.PivotTableName!,
            settings.FieldName!,
            settings.NumericGroupingStart,
            settings.NumericGroupingEnd,
            settings.NumericGroupingInterval.Value));
    }

    private int ExecuteCreateCalculatedField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "create-calculated-field"))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.Formula))
        {
            _console.WriteError("--formula is required for create-calculated-field.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.CreateCalculatedField(batch, settings.PivotTableName!, settings.FieldName!, settings.Formula));
    }

    private int ExecuteListCalculatedFields(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for list-calculated-fields.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.ListCalculatedFields(batch, settings.PivotTableName));
    }

    private int ExecuteDeleteCalculatedField(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "delete-calculated-field"))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.DeleteCalculatedField(batch, settings.PivotTableName!, settings.FieldName!));
    }

    private int ExecuteListCalculatedMembers(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for list-calculated-members.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.ListCalculatedMembers(batch, settings.PivotTableName));
    }

    private int ExecuteCreateCalculatedMember(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for create-calculated-member.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.MemberName))
        {
            _console.WriteError("--member-name is required for create-calculated-member.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.Formula))
        {
            _console.WriteError("--formula is required for create-calculated-member.");
            return -1;
        }

        if (!TryParseCalculatedMemberType(settings.MemberType, out var memberType))
        {
            return -1;
        }

        return WriteResult(_pivotTableCommands.CreateCalculatedMember(
            batch,
            settings.PivotTableName,
            settings.MemberName,
            settings.Formula,
            memberType,
            settings.SolveOrder ?? 0,
            settings.DisplayFolder,
            settings.NumberFormat));
    }

    private int ExecuteDeleteCalculatedMember(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for delete-calculated-member.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.MemberName))
        {
            _console.WriteError("--member-name is required for delete-calculated-member.");
            return -1;
        }

        return WriteResult(_pivotTableCommands.DeleteCalculatedMember(batch, settings.PivotTableName, settings.MemberName));
    }

    private bool TryParseCalculatedMemberType(string? value, out CalculatedMemberType memberType)
    {
        memberType = CalculatedMemberType.Measure;
        if (string.IsNullOrWhiteSpace(value))
        {
            return true;  // Default to Measure
        }

        if (Enum.TryParse(value, true, out memberType))
        {
            return true;
        }

        _console.WriteError("Invalid member type. Valid values: Member, Set, Measure.");
        return false;
    }

    private int ExecuteSetLayout(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for set-layout.");
            return -1;
        }

        if (!settings.LayoutType.HasValue)
        {
            _console.WriteError("--layout-type is required for set-layout. Valid values: 0=Compact, 1=Tabular, 2=Outline.");
            return -1;
        }

        if (settings.LayoutType.Value < 0 || settings.LayoutType.Value > 2)
        {
            _console.WriteError("--layout-type must be 0 (Compact), 1 (Tabular), or 2 (Outline).");
            return -1;
        }

        return WriteResult(_pivotTableCommands.SetLayout(batch, settings.PivotTableName!, settings.LayoutType.Value));
    }

    private int ExecuteSetSubtotals(IExcelBatch batch, Settings settings)
    {
        if (!ValidateFieldOperation(settings, "set-subtotals"))
        {
            return -1;
        }

        if (!settings.ShowSubtotals.HasValue)
        {
            _console.WriteError("--show-subtotals is required for set-subtotals (true/false).");
            return -1;
        }

        return WriteResult(_pivotTableCommands.SetSubtotals(batch, settings.PivotTableName!, settings.FieldName!, settings.ShowSubtotals.Value));
    }

    private int ExecuteSetGrandTotals(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName))
        {
            _console.WriteError("--pivot-name is required for set-grand-totals.");
            return -1;
        }

        var showRow = settings.ShowRowGrandTotals ?? true;
        var showColumn = settings.ShowColumnGrandTotals ?? true;

        return WriteResult(_pivotTableCommands.SetGrandTotals(batch, settings.PivotTableName!, showRow, showColumn));
    }

    private bool ValidateFieldOperation(Settings settings, string actionName)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName) || string.IsNullOrWhiteSpace(settings.FieldName))
        {
            _console.WriteError("--pivot-name and --field-name are required for " + actionName + ".");
            return false;
        }

        return true;
    }

    private bool TryParseAggregation(string? value, out AggregationFunction aggregation)
    {
        aggregation = AggregationFunction.Sum;
        if (string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        if (Enum.TryParse(value, true, out aggregation))
        {
            return true;
        }

        _console.WriteError("Invalid aggregation function. Valid values: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP.");
        return false;
    }

    private bool TryParseSortDirection(string? value, out SortDirection direction)
    {
        direction = SortDirection.Ascending;
        if (string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        if (Enum.TryParse(value, true, out direction))
        {
            return true;
        }

        _console.WriteError("Invalid sort direction. Valid values: Ascending, Descending.");
        return false;
    }

    private bool TryParseDateGroupingInterval(string? value, out DateGroupingInterval interval)
    {
        interval = DateGroupingInterval.Months;
        if (string.IsNullOrWhiteSpace(value))
        {
            _console.WriteError("--date-interval is required for group-by-date. Valid values: Days, Months, Quarters, Years.");
            return false;
        }

        if (Enum.TryParse(value, true, out interval))
        {
            return true;
        }

        _console.WriteError("Invalid date grouping interval. Valid values: Days, Months, Quarters, Years.");
        return false;
    }

    private List<string>? LoadFilterValues(Settings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.FilterValuesFile))
        {
            if (!System.IO.File.Exists(settings.FilterValuesFile))
            {
                _console.WriteError($"Filter values file '{settings.FilterValuesFile}' was not found.");
                return null;
            }

            var fileContent = System.IO.File.ReadAllText(settings.FilterValuesFile);
            var fromJson = ParseFilterValuesJson(fileContent);
            if (fromJson == null)
            {
                _console.WriteError($"Unable to parse JSON array from '{settings.FilterValuesFile}'.");
                return null;
            }

            return fromJson;
        }

        if (!string.IsNullOrWhiteSpace(settings.FilterValues))
        {
            var values = settings.FilterValues
                .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .ToList();

            if (values.Count == 0)
            {
                _console.WriteError("Provide at least one filter value.");
                return null;
            }

            return values;
        }

        _console.WriteError("Provide filter values using --filter-values or --filter-values-file.");
        return null;
    }

    private static List<string>? ParseFilterValuesJson(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var values = new List<string>();
            foreach (var element in document.RootElement.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.String)
                {
                    return null;
                }

                var value = element.GetString();
                if (!string.IsNullOrEmpty(value))
                {
                    values.Add(value);
                }
            }

            return values.Count > 0 ? values : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private bool TryGetTimeout(int? timeoutSeconds, out TimeSpan? timeout)
    {
        timeout = null;
        if (!timeoutSeconds.HasValue)
        {
            return true;
        }

        if (timeoutSeconds <= 0)
        {
            _console.WriteError("--timeout-seconds must be greater than zero when provided.");
            return false;
        }

        timeout = TimeSpan.FromSeconds(timeoutSeconds.Value);
        return true;
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown PivotTable action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--pivot-name <NAME>")]
        public string? PivotTableName { get; init; }

        [CommandOption("--sheet <SHEET>")]
        public string? SheetName { get; init; }

        [CommandOption("--range <RANGE>")]
        public string? Range { get; init; }

        [CommandOption("--destination-sheet <SHEET>")]
        public string? DestinationSheet { get; init; }

        [CommandOption("--destination-cell <CELL>")]
        public string? DestinationCell { get; init; }

        [CommandOption("--table-name <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--data-model-table <NAME>")]
        public string? DataModelTableName { get; init; }

        [CommandOption("--field-name <NAME>")]
        public string? FieldName { get; init; }

        [CommandOption("--custom-name <NAME>")]
        public string? CustomName { get; init; }

        [CommandOption("--number-format <FORMAT>")]
        public string? NumberFormat { get; init; }

        [CommandOption("--position <NUMBER>")]
        public int? Position { get; init; }

        [CommandOption("--aggregation <FUNCTION>")]
        public string? Aggregation { get; init; }

        [CommandOption("--sort-direction <DIRECTION>")]
        public string? SortDirection { get; init; }

        [CommandOption("--filter-values <CSV>")]
        public string? FilterValues { get; init; }

        [CommandOption("--filter-values-file <PATH>")]
        public string? FilterValuesFile { get; init; }

        [CommandOption("--timeout-seconds <SECONDS>")]
        public int? TimeoutSeconds { get; init; }

        [CommandOption("--date-interval <INTERVAL>")]
        public string? DateGroupingInterval { get; init; }

        [CommandOption("--numeric-interval <SIZE>")]
        public double? NumericGroupingInterval { get; init; }

        [CommandOption("--numeric-start <VALUE>")]
        public double? NumericGroupingStart { get; init; }

        [CommandOption("--numeric-end <VALUE>")]
        public double? NumericGroupingEnd { get; init; }

        [CommandOption("--formula <FORMULA>")]
        public string? Formula { get; init; }

        [CommandOption("--layout-type <TYPE>")]
        public int? LayoutType { get; init; }

        [CommandOption("--show-subtotals <BOOL>")]
        public bool? ShowSubtotals { get; init; }

        [CommandOption("--show-row-grand-totals <BOOL>")]
        public bool? ShowRowGrandTotals { get; init; }

        [CommandOption("--show-column-grand-totals <BOOL>")]
        public bool? ShowColumnGrandTotals { get; init; }

        [CommandOption("--member-name <NAME>")]
        public string? MemberName { get; init; }

        [CommandOption("--member-type <TYPE>")]
        public string? MemberType { get; init; }

        [CommandOption("--solve-order <ORDER>")]
        public int? SolveOrder { get; init; }

        [CommandOption("--display-folder <PATH>")]
        public string? DisplayFolder { get; init; }
    }
}
