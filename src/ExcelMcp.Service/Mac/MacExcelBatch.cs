using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelBatch(MacExcelBackend backend, MacExcelSession session) : NonComExcelBatch
{
    private bool _disposed;

    public override string WorkbookPath => session.FilePath;
    public override TimeSpan OperationTimeout => session.OperationTimeout;
    public override bool IsExcelVisible => session.IsVisible;
    public override bool HasTimedOutOperation => session.RequiresRecovery;

    internal static MacExcelBatch From(IExcelBatch batch) =>
        batch as MacExcelBatch ?? throw new ArgumentException("A Mac command requires a Mac workbook batch.", nameof(batch));

    internal T Invoke<T>(string command, object arguments, IReadOnlyCollection<string>? requiredHelperPrimitives = null)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var capability = MacCommandCapabilities.Get(command);
        if (!capability.IsAvailable)
        {
            throw new PlatformNotSupportedException(capability.UnavailableMessage);
        }
        var started = System.Diagnostics.Stopwatch.GetTimestamp();
        var parts = command.Split('.', 2);
        var category = parts[0];
        var action = parts[1];
        var prepared = JsonSerializer.SerializeToNode(arguments, ServiceProtocol.JsonOptions)?.AsObject()
            ?? throw new InvalidOperationException("Mac command arguments must be an object.");
        if (category == "sheet" && action is "list" or "create"
            && prepared["filePath"]?.GetValue<string>() is { Length: > 0 } path
            && !string.Equals(MacPathCanonicalizer.Normalize(path), session.FilePath, StringComparison.Ordinal))
        {
            throw new PlatformNotSupportedException(
                "macOS worksheet operations support only the exact session workbook. " +
                "Open the requested workbook in its own session; multi-workbook filePath selection requires Windows.");
        }
        PrepareArguments(category, action, prepared);
        prepared["filePath"] = session.FilePath;
        if (category == "sheet" && action is "create" or "rename")
            WorksheetCommandValidation.ValidateNewSheetName(prepared[action == "create" ? "sheetName" : "newName"]!.GetValue<string>());
        if (category == "range" && action is "get-formulas" or "set-formulas")
            return InvokeFormulas<T>(action, prepared, started);
        if (command == "calculationmode.calculate")
            return InvokeCalculation<T>(prepared, started);
        if (requiredHelperPrimitives is { Count: > 0 })
        {
            backend.RequireHelperAsync(requiredHelperPrimitives, MacAppleEvents.Remaining(OperationTimeout, started))
                .GetAwaiter().GetResult();
        }
        else if (capability.RequiredTier == MacCapabilityTier.OptionalNativeHelper)
        {
            throw new PlatformNotSupportedException("A helper-backed command must declare its required primitives before mutation.");
        }
        if (category == "sheet" && action is "create" or "rename" or "delete")
        {
            var name = prepared[action == "rename" ? "oldName" : "sheetName"]!.GetValue<string>();
            var listing = backend.InvokeAsync("sheet.list", new { filePath = session.FilePath },
                MacAppleEvents.Remaining(OperationTimeout, started)).GetAwaiter().GetResult();
            if (!listing.TryGetProperty("worksheets", out var sheetList) || sheetList.ValueKind != JsonValueKind.Array
                || sheetList.EnumerateArray().Any(sheet => sheet.ValueKind != JsonValueKind.Object
                    || !sheet.TryGetProperty("name", out var sheetName) || sheetName.ValueKind != JsonValueKind.String
                    || string.IsNullOrEmpty(sheetName.GetString())))
            {
                throw new InvalidDataException("Excel did not return valid worksheet names for command validation.");
            }
            var worksheets = ReadResult<WorksheetListResult>(listing, "sheet.list");
            if (action != "create")
                WorksheetCommandValidation.RequireExistingSheet(
                    worksheets.Worksheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.Ordinal)), name);
            if (action != "delete")
            {
                var newName = prepared[action == "create" ? "sheetName" : "newName"]!.GetValue<string>();
                foreach (var sheet in worksheets.Worksheets)
                    WorksheetCommandValidation.RequireAvailableName(sheet.Name, newName, action == "rename" ? name : null);
                backend.InvokeAsync("sheet.check-name-scope", new { filePath = session.FilePath },
                    MacAppleEvents.Remaining(OperationTimeout, started)).GetAwaiter().GetResult();
            }
        }
        var dispatchCommand = category == "worksheetstyle" ? $"sheet.{action}" : command;
        var result = backend.InvokeAsync(dispatchCommand, prepared, MacAppleEvents.Remaining(OperationTimeout, started),
            allowFailureResult: category == "pythoninexcel").GetAwaiter().GetResult();
        return ReadResult<T>(result, command);
    }

    private T InvokeCalculation<T>(JsonObject arguments, long started)
    {
        TimeSpan Remaining() => MacAppleEvents.Remaining(OperationTimeout, started);
        var scope = arguments["scope"]!.Deserialize<CalculationScope>(ServiceProtocol.JsonOptions);
        var kind = arguments["kind"]?.Deserialize<CalculationKind>(ServiceProtocol.JsonOptions) ?? CalculationKind.Normal;
        var name = arguments["sheetName"]?.GetValue<string>();
        var address = arguments["rangeAddress"]?.GetValue<string>();
        CalculationCommandValidation.Validate(scope, name, address, kind);
        if (scope == CalculationScope.Application)
            throw new PlatformNotSupportedException(
                "macOS application-scope calculation would affect unrelated workbooks in shared Excel; no calculation was attempted.");
        var listing = ReadResult<WorksheetListResult>(
            backend.InvokeAsync("sheet.list", new { filePath = session.FilePath }, Remaining()).GetAwaiter().GetResult(), "sheet.list");
        if (!listing.Worksheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.Ordinal)))
            throw new InvalidOperationException($"Worksheet '{name}' was not found.");
        if (scope == CalculationScope.Range)
        {
            var areas = RangeHelpers.ParseSupportedRangeAreas(address!);
            if (areas is null)
                throw new OperationFailureException(OperationFailureCategory.InvalidInput,
                    $"Sheet '{name}' exists, but range '{address}' is invalid. Verify the range address format (e.g., 'A1:E10', 'A1', 'A:A').");
            if (areas.Count != 1 || areas[0].Contains('[', StringComparison.Ordinal) || areas[0].EndsWith('#'))
                throw new PlatformNotSupportedException(
                    "Native macOS calculation does not yet support disjoint, structured-reference, or spill-address variants.");
            arguments["rangeAddress"] = areas[0];
        }
        return ReadResult<T>(backend.InvokeAsync("calculation.calculate", arguments, Remaining()).GetAwaiter().GetResult(),
            "calculation.calculate");
    }

    private T InvokeFormulas<T>(string action, JsonObject arguments, long started)
    {
        TimeSpan Remaining() => MacAppleEvents.Remaining(OperationTimeout, started);
        var name = arguments["sheetName"]!.GetValue<string>();
        var address = arguments["rangeAddress"]!.GetValue<string>();
        if (string.IsNullOrEmpty(name))
            throw new PlatformNotSupportedException("Native macOS formula actions do not yet support named-range variants.");
        var listing = backend.InvokeAsync("sheet.list", new { filePath = session.FilePath }, Remaining()).GetAwaiter().GetResult();
        var worksheets = ReadResult<WorksheetListResult>(listing, "sheet.list");
        if (!worksheets.Worksheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.Ordinal)))
            throw new OperationFailureException(OperationFailureCategory.NotFound, $"Sheet '{name}' not found.");
        var areas = RangeHelpers.ParseSupportedRangeAreas(address);
        if (areas is null)
            throw new OperationFailureException(OperationFailureCategory.InvalidInput,
                $"Sheet '{name}' exists, but range '{address}' is invalid. Verify the range address format (e.g., 'A1:E10', 'A1', 'A:A').");
        if (areas.Count != 1 || areas[0].Contains('[', StringComparison.Ordinal) || areas[0].EndsWith('#'))
            throw new PlatformNotSupportedException(
                "Native macOS formula actions do not yet support disjoint, structured-reference, or spill-address variants.");
        arguments["rangeAddress"] = areas[0];
        arguments["forWrite"] = action == "set-formulas";
        var description = backend.InvokeAsync("range.describe", arguments, Remaining()).GetAwaiter().GetResult();
        var geometry = ReadResult<MacRangeGeometry>(description, "range.describe");
        if (geometry.Rows <= 0 || geometry.Columns <= 0 || geometry.Row <= 0 || geometry.Column <= 0
            || string.IsNullOrEmpty(geometry.Address))
            throw new InvalidDataException("Excel returned invalid native range geometry.");
        if (action == "set-formulas")
        {
            if (geometry.MergedRanges.Count > 0 && !geometry.IsMergedTopLeft)
                RangeCommandValidation.ThrowMergedCellWriteError(address, geometry.MergedRanges);
            var formulas = arguments["formulas"]!.Deserialize<List<List<string>>>(ServiceProtocol.JsonOptions)!;
            RangeCommandValidation.ValidateDimensions(formulas, geometry.Rows, geometry.Columns, "formulas", "Formula");
            var policy = arguments["overwritePolicy"]?.Deserialize<OverwritePolicy>(ServiceProtocol.JsonOptions) ?? OverwritePolicy.RejectNonempty;
            if (policy != OverwritePolicy.Allow)
            {
                var data = ReadNativeData(arguments, Remaining());
                var conflicts = new List<string>();
                for (var row = 0; row < geometry.Rows; row++)
                    for (var column = 0; column < geometry.Columns; column++)
                    {
                        if (!RangeCommandValidation.IsOccupied(MacNativeRange.Scalar(data.Values[row]![column]),
                            MacNativeRange.Scalar(data.Formulas[row]![column])))
                            continue;
                        if (conflicts.Count == RangeCommandValidation.ConflictExampleLimit)
                            RangeCommandValidation.ThrowOccupiedDestination(name, conflicts, true);
                        conflicts.Add($"${RangeCommandValidation.ColumnLetter(geometry.Column + column)}${geometry.Row + row}");
                    }
                if (conflicts.Count > 0) RangeCommandValidation.ThrowOccupiedDestination(name, conflicts, false);
            }
            return ReadResult<T>(backend.InvokeAsync("range.set-formulas", arguments, Remaining()).GetAwaiter().GetResult(), "range.set-formulas");
        }
        var read = ReadNativeData(arguments, Remaining());
        var formulaArray = new object[read.Rows, read.Columns];
        var valueArray = new object[read.Rows, read.Columns];
        for (var row = 0; row < read.Rows; row++)
            for (var column = 0; column < read.Columns; column++)
            {
                formulaArray[row, column] = MacNativeRange.Scalar(read.Formulas[row]![column])!;
                valueArray[row, column] = MacNativeRange.Scalar(read.Values[row]![column])!;
            }
        var result = RangeFormulaResults.Create(session.FilePath, name, geometry.Address, geometry.Row, geometry.Column,
            formulaArray, valueArray);
        return result is T typed ? typed : throw new InvalidOperationException("A formula read requires the shared formula result type.");
    }

    private MacRangeData ReadNativeData(JsonObject arguments, TimeSpan timeout)
    {
        var result = ReadResult<MacRangeData>(
            backend.InvokeAsync("range.read-data", arguments, timeout).GetAwaiter().GetResult(), "range.read-data");
        MacNativeRange.Matrix(result.Formulas, result.Rows, result.Columns);
        MacNativeRange.Matrix(result.Values, result.Rows, result.Columns);
        return result;
    }

    private static T ReadResult<T>(JsonElement result, string command)
    {
        if (!result.TryGetProperty("success", out var success) || success.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
        {
            throw new InvalidDataException($"Excel did not return a success state for '{command}'.");
        }
        if (success.GetBoolean() && result.TryGetProperty("errorMessage", out var error)
            && error.ValueKind != JsonValueKind.Null && !string.IsNullOrEmpty(error.GetString()))
        {
            throw new InvalidDataException($"Excel returned an inconsistent success state for '{command}'.");
        }
        return result.Deserialize<T>(ServiceProtocol.JsonOptions)
            ?? throw new InvalidDataException($"Excel did not return the required result for '{command}'.");
    }

    internal void Invoke(string command, object arguments) => Invoke<OperationResult>(command, arguments);

    private void PrepareArguments(string category, string action, JsonObject arguments)
    {
        if (category == "table" && action == "append")
        {
            var rows = arguments["rows"]?.Deserialize<List<List<object?>>>(ServiceProtocol.JsonOptions);
            var rowsFile = arguments["rowsFile"]?.GetValue<string>();
            arguments["rows"] = JsonSerializer.SerializeToNode(
                Core.Utilities.ParameterTransforms.ResolveValuesOrFile(rows, rowsFile, "rows"), ServiceProtocol.JsonOptions);
            arguments.Remove("rowsFile");
        }
        MacRangeArguments.Prepare(category, action, arguments);
        if (category == "pythoninexcel") MacPythonInExcelArguments.Prepare(action, arguments, OperationTimeout);
        if (category == "namedrange") MacNamedRangeArguments.Prepare(action, arguments);
        if (category == "range")
        {
            var nameCommand = action == "set-values" ? "namedrange.write" : "namedrange.read";
            MacNamedRangeArguments.PrepareRangeBinding(action, arguments, MacCommandCapabilities.Get(nameCommand).IsAvailable);
        }
        if (category == "rangeformat" && action == "set-column-width")
        {
            var width = arguments["columnWidth"]!.GetValue<double>();
            if (width is < 0.25 or > 409) throw new ArgumentException("columnWidth must be between 0.25 and 409 points");
        }
        if (category == "rangeformat" && action == "set-row-height")
        {
            var height = arguments["rowHeight"]!.GetValue<double>();
            if (height is < 0 or > 409) throw new ArgumentException("rowHeight must be between 0 and 409 points");
        }
    }

    public override void UpdateWorkbookPath(string workbookPath) => throw new PlatformNotSupportedException("Mac Save As requires verified workbook identity and path reservation support.");
    public override void Save(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Invoke("workbook.save", new { });
    }
    public override void Dispose() => _disposed = true;
}
