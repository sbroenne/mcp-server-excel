using System.Globalization;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Range;

internal sealed class RangeCommand : Command<RangeCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IRangeCommands _rangeCommands;
    private readonly ICliConsole _console;

    public RangeCommand(ISessionService sessionService, IRangeCommands rangeCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _rangeCommands = rangeCommands ?? throw new ArgumentNullException(nameof(rangeCommands));
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
            "get-values" => ExecuteGetValues(batch, settings),
            "set-values" => ExecuteSetValues(batch, settings),
            "get-formulas" => ExecuteGetFormulas(batch, settings),
            "set-formulas" => ExecuteSetFormulas(batch, settings),
            "clear-all" => ExecuteClear(batch, settings, ClearType.All),
            "clear-contents" => ExecuteClear(batch, settings, ClearType.Contents),
            "clear-formats" => ExecuteClear(batch, settings, ClearType.Formats),
            "copy" => ExecuteCopy(batch, settings, CopyType.All),
            "copy-values" => ExecuteCopy(batch, settings, CopyType.Values),
            "copy-formulas" => ExecuteCopy(batch, settings, CopyType.Formulas),
            "insert-cells" => ExecuteInsertCells(batch, settings),
            "delete-cells" => ExecuteDeleteCells(batch, settings),
            "insert-rows" => ExecuteInsertRows(batch, settings),
            "delete-rows" => ExecuteDeleteRows(batch, settings),
            "insert-columns" => ExecuteInsertColumns(batch, settings),
            "delete-columns" => ExecuteDeleteColumns(batch, settings),
            "find" => ExecuteFind(batch, settings),
            "replace" => ExecuteReplace(batch, settings),
            "sort" => ExecuteSort(batch, settings),
            "get-used-range" => ExecuteGetUsedRange(batch, settings),
            "get-current-region" => ExecuteGetCurrentRegion(batch, settings),
            "get-info" => ExecuteGetInfo(batch, settings),
            "add-hyperlink" => ExecuteAddHyperlink(batch, settings),
            "remove-hyperlink" => ExecuteRemoveHyperlink(batch, settings),
            "list-hyperlinks" => ExecuteListHyperlinks(batch, settings),
            "get-hyperlink" => ExecuteGetHyperlink(batch, settings),
            "get-number-formats" => ExecuteGetNumberFormats(batch, settings),
            "set-number-format" => ExecuteSetNumberFormat(batch, settings),
            "set-number-formats" => ExecuteSetNumberFormats(batch, settings),
            "get-style" => ExecuteGetStyle(batch, settings),
            "set-style" => ExecuteSetStyle(batch, settings),
            "format-range" => ExecuteFormatRange(batch, settings),
            "validate-range" => ExecuteValidateRange(batch, settings),
            "get-validation" => ExecuteGetValidation(batch, settings),
            "remove-validation" => ExecuteRemoveValidation(batch, settings),
            "autofit-columns" => ExecuteAutoFitColumns(batch, settings),
            "autofit-rows" => ExecuteAutoFitRows(batch, settings),
            "merge-cells" => ExecuteMergeCells(batch, settings),
            "unmerge-cells" => ExecuteUnmergeCells(batch, settings),
            "get-merge-info" => ExecuteGetMergeInfo(batch, settings),
            "set-cell-lock" => ExecuteSetCellLock(batch, settings),
            "get-cell-lock" => ExecuteGetCellLock(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteGetValues(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-values.");
            return -1;
        }

        var result = _rangeCommands.GetValues(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteSetValues(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for set-values.");
            return -1;
        }

        var values = LoadValues(settings, _console);
        if (values == null)
        {
            return -1;
        }

        var result = _rangeCommands.SetValues(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, values);
        return WriteResult(result);
    }

    private int ExecuteGetFormulas(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-formulas.");
            return -1;
        }

        var result = _rangeCommands.GetFormulas(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteSetFormulas(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for set-formulas.");
            return -1;
        }

        var formulas = LoadFormulas(settings, _console);
        if (formulas == null)
        {
            return -1;
        }

        var result = _rangeCommands.SetFormulas(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, formulas);
        return WriteResult(result);
    }

    private int ExecuteClear(IExcelBatch batch, Settings settings, ClearType type)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for clear operations.");
            return -1;
        }

        var sheet = settings.SheetName ?? string.Empty;
        OperationResult result = type switch
        {
            ClearType.All => _rangeCommands.ClearAll(batch, sheet, settings.RangeAddress),
            ClearType.Contents => _rangeCommands.ClearContents(batch, sheet, settings.RangeAddress),
            ClearType.Formats => _rangeCommands.ClearFormats(batch, sheet, settings.RangeAddress),
            _ => throw new InvalidOperationException("Unsupported clear type.")
        };

        return WriteResult(result);
    }

    private int ExecuteCopy(IExcelBatch batch, Settings settings, CopyType type)
    {
        if (string.IsNullOrWhiteSpace(settings.SourceSheet) || string.IsNullOrWhiteSpace(settings.SourceRange) ||
            string.IsNullOrWhiteSpace(settings.TargetSheet) || string.IsNullOrWhiteSpace(settings.TargetRange))
        {
            _console.WriteError("Source and target sheet/range are required for copy operations.");
            return -1;
        }

        OperationResult result = type switch
        {
            CopyType.All => _rangeCommands.Copy(batch, settings.SourceSheet, settings.SourceRange, settings.TargetSheet, settings.TargetRange),
            CopyType.Values => _rangeCommands.CopyValues(batch, settings.SourceSheet, settings.SourceRange, settings.TargetSheet, settings.TargetRange),
            CopyType.Formulas => _rangeCommands.CopyFormulas(batch, settings.SourceSheet, settings.SourceRange, settings.TargetSheet, settings.TargetRange),
            _ => throw new InvalidOperationException("Unsupported copy type.")
        };

        return WriteResult(result);
    }

    private int ExecuteAddHyperlink(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.CellAddress) || string.IsNullOrWhiteSpace(settings.Url) || string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("Sheet, cell, and URL are required for add-hyperlink.");
            return -1;
        }

        var result = _rangeCommands.AddHyperlink(batch, settings.SheetName, settings.CellAddress, settings.Url, settings.DisplayText, settings.Tooltip);
        return WriteResult(result);
    }

    private int ExecuteRemoveHyperlink(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress) || string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("Sheet and range are required for remove-hyperlink.");
            return -1;
        }

        var result = _rangeCommands.RemoveHyperlink(batch, settings.SheetName, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteListHyperlinks(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("Sheet is required for list-hyperlinks.");
            return -1;
        }

        var result = _rangeCommands.ListHyperlinks(batch, settings.SheetName);
        return WriteResult(result);
    }

    private int ExecuteGetHyperlink(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.CellAddress) || string.IsNullOrWhiteSpace(settings.SheetName))
        {
            _console.WriteError("Sheet and cell are required for get-hyperlink.");
            return -1;
        }

        var result = _rangeCommands.GetHyperlink(batch, settings.SheetName, settings.CellAddress);
        return WriteResult(result);
    }

    // === INSERT/DELETE OPERATIONS ===

    private int ExecuteInsertCells(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for insert-cells.");
            return -1;
        }

        if (!TryParseInsertShiftDirection(settings.Shift, out var shift))
        {
            _console.WriteError($"Invalid shift direction '{settings.Shift}'. Use 'down' or 'right'.");
            return -1;
        }

        var result = _rangeCommands.InsertCells(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, shift);
        return WriteResult(result);
    }

    private int ExecuteDeleteCells(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for delete-cells.");
            return -1;
        }

        if (!TryParseDeleteShiftDirection(settings.Shift, out var shift))
        {
            _console.WriteError($"Invalid shift direction '{settings.Shift}'. Use 'up' or 'left'.");
            return -1;
        }

        var result = _rangeCommands.DeleteCells(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, shift);
        return WriteResult(result);
    }

    private int ExecuteInsertRows(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for insert-rows.");
            return -1;
        }

        var result = _rangeCommands.InsertRows(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteDeleteRows(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for delete-rows.");
            return -1;
        }

        var result = _rangeCommands.DeleteRows(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteInsertColumns(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for insert-columns.");
            return -1;
        }

        var result = _rangeCommands.InsertColumns(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteDeleteColumns(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for delete-columns.");
            return -1;
        }

        var result = _rangeCommands.DeleteColumns(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    // === FIND/REPLACE OPERATIONS ===

    private int ExecuteFind(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for find.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.SearchValue))
        {
            _console.WriteError("Search value is required for find.");
            return -1;
        }

        var options = new FindOptions
        {
            MatchCase = settings.MatchCase ?? false,
            MatchEntireCell = settings.MatchEntireCell ?? false,
            SearchFormulas = settings.SearchFormulas ?? true,
            SearchValues = settings.SearchValues ?? true,
            SearchComments = settings.SearchComments ?? false
        };

        var result = _rangeCommands.Find(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, settings.SearchValue, options);
        return WriteResult(result);
    }

    private int ExecuteReplace(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for replace.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.FindValue))
        {
            _console.WriteError("Find value is required for replace.");
            return -1;
        }

        if (settings.ReplaceValue == null)
        {
            _console.WriteError("Replace value is required for replace.");
            return -1;
        }

        var options = new ReplaceOptions
        {
            ReplaceAll = settings.ReplaceAll ?? true,
            MatchCase = settings.MatchCase ?? false,
            MatchEntireCell = settings.MatchEntireCell ?? false,
            SearchFormulas = settings.SearchFormulas ?? true,
            SearchValues = settings.SearchValues ?? true,
            SearchComments = settings.SearchComments ?? false
        };

        _rangeCommands.Replace(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, settings.FindValue, settings.ReplaceValue, options);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    // === SORT OPERATION ===

    private int ExecuteSort(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for sort.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.SortColumnsJson))
        {
            _console.WriteError("Sort columns JSON is required for sort. Example: [{\"columnIndex\":1,\"ascending\":true}]");
            return -1;
        }

        List<SortColumn>? sortColumns;
        try
        {
            sortColumns = JsonSerializer.Deserialize<List<SortColumn>>(settings.SortColumnsJson);
            if (sortColumns == null || sortColumns.Count == 0)
            {
                _console.WriteError("Sort columns JSON must contain at least one column.");
                return -1;
            }
        }
        catch (JsonException ex)
        {
            _console.WriteError($"Invalid sort columns JSON: {ex.Message}");
            return -1;
        }

        _rangeCommands.Sort(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, sortColumns, settings.HasHeaders ?? true);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    // === RANGE QUERIES ===

    private int ExecuteGetUsedRange(IExcelBatch batch, Settings settings)
    {
        var result = _rangeCommands.GetUsedRange(batch, settings.SheetName ?? string.Empty);
        return WriteResult(result);
    }

    private int ExecuteGetCurrentRegion(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.CellAddress))
        {
            _console.WriteError("Cell address is required for get-current-region.");
            return -1;
        }

        var result = _rangeCommands.GetCurrentRegion(batch, settings.SheetName ?? string.Empty, settings.CellAddress);
        return WriteResult(result);
    }

    private int ExecuteGetInfo(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-info.");
            return -1;
        }

        var result = _rangeCommands.GetInfo(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    // === NUMBER FORMAT OPERATIONS ===

    private int ExecuteGetNumberFormats(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-number-formats.");
            return -1;
        }

        var result = _rangeCommands.GetNumberFormats(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteSetNumberFormat(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for set-number-format.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.FormatCode))
        {
            _console.WriteError("Format code is required for set-number-format.");
            return -1;
        }

        var result = _rangeCommands.SetNumberFormat(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, settings.FormatCode);
        return WriteResult(result);
    }

    private int ExecuteSetNumberFormats(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for set-number-formats.");
            return -1;
        }

        var formats = LoadFormats(settings, _console);
        if (formats == null)
        {
            return -1;
        }

        var result = _rangeCommands.SetNumberFormats(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, formats);
        return WriteResult(result);
    }

    // === STYLE OPERATIONS ===

    private int ExecuteGetStyle(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-style.");
            return -1;
        }

        var result = _rangeCommands.GetStyle(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteSetStyle(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for set-style.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.StyleName))
        {
            _console.WriteError("Style name is required for set-style.");
            return -1;
        }

        _rangeCommands.SetStyle(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, settings.StyleName);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    // === FORMATTING OPERATION ===

    private int ExecuteFormatRange(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for format-range.");
            return -1;
        }

        _rangeCommands.FormatRange(
            batch,
            settings.SheetName ?? string.Empty,
            settings.RangeAddress,
            settings.FontName,
            settings.FontSize,
            settings.Bold,
            settings.Italic,
            settings.Underline,
            settings.FontColor,
            settings.FillColor,
            settings.BorderStyle,
            settings.BorderColor,
            settings.BorderWeight,
            settings.HorizontalAlignment,
            settings.VerticalAlignment,
            settings.WrapText,
            settings.Orientation);

        _console.WriteJson(new { Success = true });
        return 0;
    }

    // === VALIDATION OPERATIONS ===

    private int ExecuteValidateRange(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for validate-range.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.ValidationType))
        {
            _console.WriteError("Validation type is required for validate-range.");
            return -1;
        }

        _rangeCommands.ValidateRange(
            batch,
            settings.SheetName ?? string.Empty,
            settings.RangeAddress,
            settings.ValidationType,
            settings.ValidationOperator,
            settings.ValidationFormula1,
            settings.ValidationFormula2,
            settings.ShowInputMessage,
            settings.InputTitle,
            settings.InputMessage,
            settings.ShowErrorAlert,
            settings.ErrorStyle,
            settings.ErrorTitle,
            settings.ErrorMessage,
            settings.IgnoreBlank,
            settings.ShowDropdown);

        _console.WriteJson(new { Success = true });
        return 0;
    }

    private int ExecuteGetValidation(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-validation.");
            return -1;
        }

        var result = _rangeCommands.GetValidation(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    private int ExecuteRemoveValidation(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for remove-validation.");
            return -1;
        }

        _rangeCommands.RemoveValidation(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    // === AUTO-FIT OPERATIONS ===

    private int ExecuteAutoFitColumns(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for autofit-columns.");
            return -1;
        }

        _rangeCommands.AutoFitColumns(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    private int ExecuteAutoFitRows(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for autofit-rows.");
            return -1;
        }

        _rangeCommands.AutoFitRows(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    // === MERGE OPERATIONS ===

    private int ExecuteMergeCells(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for merge-cells.");
            return -1;
        }

        _rangeCommands.MergeCells(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    private int ExecuteUnmergeCells(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for unmerge-cells.");
            return -1;
        }

        _rangeCommands.UnmergeCells(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    private int ExecuteGetMergeInfo(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-merge-info.");
            return -1;
        }

        var result = _rangeCommands.GetMergeInfo(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    // === CELL LOCK OPERATIONS ===

    private int ExecuteSetCellLock(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for set-cell-lock.");
            return -1;
        }

        if (!settings.Locked.HasValue)
        {
            _console.WriteError("Locked flag is required for set-cell-lock.");
            return -1;
        }

        _rangeCommands.SetCellLock(batch, settings.SheetName ?? string.Empty, settings.RangeAddress, settings.Locked.Value);
        _console.WriteJson(new { Success = true });
        return 0;
    }

    private int ExecuteGetCellLock(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.RangeAddress))
        {
            _console.WriteError("Range address is required for get-cell-lock.");
            return -1;
        }

        var result = _rangeCommands.GetCellLock(batch, settings.SheetName ?? string.Empty, settings.RangeAddress);
        return WriteResult(result);
    }

    // === HELPER METHODS ===

    private static bool TryParseInsertShiftDirection(string? direction, out InsertShiftDirection result)
    {
        result = InsertShiftDirection.Down;
        if (string.IsNullOrWhiteSpace(direction))
        {
            return true; // Default to Down
        }

        return direction.ToLowerInvariant() switch
        {
            "down" => SetResult(out result, InsertShiftDirection.Down),
            "right" => SetResult(out result, InsertShiftDirection.Right),
            _ => false
        };
    }

    private static bool TryParseDeleteShiftDirection(string? direction, out DeleteShiftDirection result)
    {
        result = DeleteShiftDirection.Up;
        if (string.IsNullOrWhiteSpace(direction))
        {
            return true; // Default to Up
        }

        return direction.ToLowerInvariant() switch
        {
            "up" => SetResult(out result, DeleteShiftDirection.Up),
            "left" => SetResult(out result, DeleteShiftDirection.Left),
            _ => false
        };
    }

    private static bool SetResult<T>(out T result, T value)
    {
        result = value;
        return true;
    }

    private static List<List<string>>? LoadFormats(Settings settings, ICliConsole console)
    {
        if (!string.IsNullOrWhiteSpace(settings.FormatsJson))
        {
            var formats = ParseFormatsJson(settings.FormatsJson!);
            if (formats == null)
            {
                console.WriteError("Unable to parse --formats-json content.");
            }

            return formats;
        }

        console.WriteError("Provide --formats-json for set-number-formats.");
        return null;
    }

    private static List<List<string>>? ParseFormatsJson(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var result = new List<List<string>>();
            foreach (var row in document.RootElement.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                var rowList = new List<string>();
                foreach (var cell in row.EnumerateArray())
                {
                    rowList.Add(cell.ValueKind == JsonValueKind.String ? cell.GetString() ?? string.Empty : string.Empty);
                }

                result.Add(rowList);
            }

            return result;
        }
        catch
        {
            return null;
        }
    }

    private static List<List<object?>>? LoadValues(Settings settings, ICliConsole console)
    {
        if (!string.IsNullOrWhiteSpace(settings.ValuesJson))
        {
            var values = ParseValuesJson(settings.ValuesJson!);
            if (values == null)
            {
                console.WriteError("Unable to parse --values-json content.");
            }

            return values;
        }

        if (!string.IsNullOrWhiteSpace(settings.ValuesCsv))
        {
            var values = ParseValuesCsv(settings.ValuesCsv!);
            if (values == null)
            {
                console.WriteError($"Unable to parse CSV file '{settings.ValuesCsv}'.");
            }

            return values;
        }

        console.WriteError("Provide --values-json or --values-csv for set-values.");
        return null;
    }

    private static List<List<string>>? LoadFormulas(Settings settings, ICliConsole console)
    {
        if (!string.IsNullOrWhiteSpace(settings.FormulasJson))
        {
            var formulas = ParseFormulasJson(settings.FormulasJson!);
            if (formulas == null)
            {
                console.WriteError("Unable to parse --formulas-json content.");
            }

            return formulas;
        }

        if (!string.IsNullOrWhiteSpace(settings.FormulasCsv))
        {
            var formulas = ParseFormulasCsv(settings.FormulasCsv!);
            if (formulas == null)
            {
                console.WriteError($"Unable to parse CSV file '{settings.FormulasCsv}'.");
            }

            return formulas;
        }

        console.WriteError("Provide --formulas-json or --formulas-csv for set-formulas.");
        return null;
    }

    private static List<List<object?>>? ParseValuesJson(string json)
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

    private static List<List<string>>? ParseFormulasJson(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<List<string>>();
            foreach (var rowElement in document.RootElement.EnumerateArray())
            {
                if (rowElement.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                var row = new List<string>();
                foreach (var cell in rowElement.EnumerateArray())
                {
                    row.Add(cell.ValueKind == JsonValueKind.String ? cell.GetString() ?? string.Empty : cell.GetRawText());
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

    private static List<List<object?>>? ParseValuesCsv(string path)
    {
        var rows = ParseCsv(path);
        if (rows == null)
        {
            return null;
        }

        return rows.Select(row => row.Select(ConvertScalar).ToList()).ToList();
    }

    private static List<List<string>>? ParseFormulasCsv(string path)
    {
        var rows = ParseCsv(path);
        return rows?.Select(row => row.Select(value => value ?? string.Empty).ToList()).ToList();
    }

    private static List<List<string?>>? ParseCsv(string path)
    {
        if (!System.IO.File.Exists(path))
        {
            return null;
        }

        var lines = System.IO.File.ReadAllLines(path);
        var rows = new List<List<string?>>();

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            rows.Add(ParseCsvLine(line));
        }

        return rows;
    }

    private static List<string?> ParseCsvLine(string line)
    {
        var values = new List<string?>();
        var builder = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var character = line[i];

            if (character == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    builder.Append('"');
                    i++;
                    continue;
                }

                inQuotes = !inQuotes;
                continue;
            }

            if (character == ',' && !inQuotes)
            {
                values.Add(builder.Length == 0 ? null : builder.ToString().Trim());
                builder.Clear();
                continue;
            }

            builder.Append(character);
        }

        values.Add(builder.Length == 0 ? null : builder.ToString().Trim());
        return values;
    }

    private static object? ConvertScalar(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (bool.TryParse(value, out var boolValue))
        {
            return boolValue;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var invariantNumber))
        {
            return invariantNumber;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out var cultureNumber))
        {
            return cultureNumber;
        }

        return value;
    }

    private static object? ConvertJsonCell(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.Null => null,
            JsonValueKind.Undefined => null,
            JsonValueKind.Number => element.TryGetInt64(out var i64) ? i64 : element.GetDouble(),
            JsonValueKind.String => element.GetString(),
            JsonValueKind.False => false,
            JsonValueKind.True => true,
            _ => element.GetRawText()
        };
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown range action '{action}'.");
        return -1;
    }

    private enum ClearType
    {
        All,
        Contents,
        Formats
    }

    private enum CopyType
    {
        All,
        Values,
        Formulas
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <SHEET>")]
        public string? SheetName { get; init; }

        [CommandOption("--range <RANGE>")]
        public string? RangeAddress { get; init; }

        [CommandOption("--source-sheet <SHEET>")]
        public string? SourceSheet { get; init; }

        [CommandOption("--source-range <RANGE>")]
        public string? SourceRange { get; init; }

        [CommandOption("--target-sheet <SHEET>")]
        public string? TargetSheet { get; init; }

        [CommandOption("--target-range <RANGE>")]
        public string? TargetRange { get; init; }

        [CommandOption("--cell <CELL>")]
        public string? CellAddress { get; init; }

        [CommandOption("--url <URL>")]
        public string? Url { get; init; }

        [CommandOption("--display-text <TEXT>")]
        public string? DisplayText { get; init; }

        [CommandOption("--tooltip <TEXT>")]
        public string? Tooltip { get; init; }

        [CommandOption("--values-json <JSON>")]
        public string? ValuesJson { get; init; }

        [CommandOption("--values-csv <PATH>")]
        public string? ValuesCsv { get; init; }

        [CommandOption("--formulas-json <JSON>")]
        public string? FormulasJson { get; init; }

        [CommandOption("--formulas-csv <PATH>")]
        public string? FormulasCsv { get; init; }

        // Insert/Delete operations
        [CommandOption("--shift <DIRECTION>")]
        public string? Shift { get; init; }

        // Find/Replace operations
        [CommandOption("--search-value <VALUE>")]
        public string? SearchValue { get; init; }

        [CommandOption("--find-value <VALUE>")]
        public string? FindValue { get; init; }

        [CommandOption("--replace-value <VALUE>")]
        public string? ReplaceValue { get; init; }

        [CommandOption("--match-case")]
        public bool? MatchCase { get; init; }

        [CommandOption("--match-entire-cell")]
        public bool? MatchEntireCell { get; init; }

        [CommandOption("--search-formulas")]
        public bool? SearchFormulas { get; init; }

        [CommandOption("--search-values")]
        public bool? SearchValues { get; init; }

        [CommandOption("--search-comments")]
        public bool? SearchComments { get; init; }

        [CommandOption("--replace-all")]
        public bool? ReplaceAll { get; init; }

        // Sort operation
        [CommandOption("--sort-columns-json <JSON>")]
        public string? SortColumnsJson { get; init; }

        [CommandOption("--has-headers")]
        public bool? HasHeaders { get; init; }

        // Number format operations
        [CommandOption("--format-code <CODE>")]
        public string? FormatCode { get; init; }

        [CommandOption("--formats-json <JSON>")]
        public string? FormatsJson { get; init; }

        // Style operations
        [CommandOption("--style-name <NAME>")]
        public string? StyleName { get; init; }

        // Format-range operation
        [CommandOption("--font-name <NAME>")]
        public string? FontName { get; init; }

        [CommandOption("--font-size <SIZE>")]
        public double? FontSize { get; init; }

        [CommandOption("--bold")]
        public bool? Bold { get; init; }

        [CommandOption("--italic")]
        public bool? Italic { get; init; }

        [CommandOption("--underline")]
        public bool? Underline { get; init; }

        [CommandOption("--font-color <COLOR>")]
        public string? FontColor { get; init; }

        [CommandOption("--fill-color <COLOR>")]
        public string? FillColor { get; init; }

        [CommandOption("--border-style <STYLE>")]
        public string? BorderStyle { get; init; }

        [CommandOption("--border-color <COLOR>")]
        public string? BorderColor { get; init; }

        [CommandOption("--border-weight <WEIGHT>")]
        public string? BorderWeight { get; init; }

        [CommandOption("--horizontal-alignment <ALIGNMENT>")]
        public string? HorizontalAlignment { get; init; }

        [CommandOption("--vertical-alignment <ALIGNMENT>")]
        public string? VerticalAlignment { get; init; }

        [CommandOption("--wrap-text")]
        public bool? WrapText { get; init; }

        [CommandOption("--orientation <DEGREES>")]
        public int? Orientation { get; init; }

        // Validation operations
        [CommandOption("--validation-type <TYPE>")]
        public string? ValidationType { get; init; }

        [CommandOption("--validation-operator <OPERATOR>")]
        public string? ValidationOperator { get; init; }

        [CommandOption("--validation-formula1 <FORMULA>")]
        public string? ValidationFormula1 { get; init; }

        [CommandOption("--validation-formula2 <FORMULA>")]
        public string? ValidationFormula2 { get; init; }

        [CommandOption("--show-input-message")]
        public bool? ShowInputMessage { get; init; }

        [CommandOption("--input-title <TITLE>")]
        public string? InputTitle { get; init; }

        [CommandOption("--input-message <MESSAGE>")]
        public string? InputMessage { get; init; }

        [CommandOption("--show-error-alert")]
        public bool? ShowErrorAlert { get; init; }

        [CommandOption("--error-style <STYLE>")]
        public string? ErrorStyle { get; init; }

        [CommandOption("--error-title <TITLE>")]
        public string? ErrorTitle { get; init; }

        [CommandOption("--error-message <MESSAGE>")]
        public string? ErrorMessage { get; init; }

        [CommandOption("--ignore-blank")]
        public bool? IgnoreBlank { get; init; }

        [CommandOption("--show-dropdown")]
        public bool? ShowDropdown { get; init; }

        // Cell lock operations
        [CommandOption("--locked")]
        public bool? Locked { get; init; }
    }
}
