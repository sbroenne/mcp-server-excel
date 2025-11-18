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
            "add-hyperlink" => ExecuteAddHyperlink(batch, settings),
            "remove-hyperlink" => ExecuteRemoveHyperlink(batch, settings),
            "list-hyperlinks" => ExecuteListHyperlinks(batch, settings),
            "get-hyperlink" => ExecuteGetHyperlink(batch, settings),
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
    }
}
