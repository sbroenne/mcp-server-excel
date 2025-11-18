using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.ConditionalFormatting;

internal sealed class ConditionalFormattingCommand : Command<ConditionalFormattingCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IConditionalFormattingCommands _formattingCommands;
    private readonly ICliConsole _console;

    public ConditionalFormattingCommand(ISessionService sessionService, IConditionalFormattingCommands formattingCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _formattingCommands = formattingCommands ?? throw new ArgumentNullException(nameof(formattingCommands));
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
            "add-rule" => ExecuteAddRule(batch, settings),
            "clear-rules" => ExecuteClearRules(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteAddRule(IExcelBatch batch, Settings settings)
    {
        if (!TryGetRangeInputs(settings, out var sheetName, out var rangeAddress))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.RuleType))
        {
            _console.WriteError("--rule-type is required for add-rule.");
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.Formula1))
        {
            _console.WriteError("--formula1 is required for add-rule.");
            return -1;
        }

        if (RequiresSecondFormula(settings.OperatorType) && string.IsNullOrWhiteSpace(settings.Formula2))
        {
            _console.WriteError("--formula2 is required when --operator is 'between' or 'notBetween'.");
            return -1;
        }

        return WriteResult(_formattingCommands.AddRule(
            batch,
            sheetName,
            rangeAddress,
            settings.RuleType!,
            settings.OperatorType,
            settings.Formula1,
            settings.Formula2,
            settings.InteriorColor,
            settings.InteriorPattern,
            settings.FontColor,
            settings.FontBold,
            settings.FontItalic,
            settings.BorderStyle,
            settings.BorderColor));
    }

    private int ExecuteClearRules(IExcelBatch batch, Settings settings)
    {
        if (!TryGetRangeInputs(settings, out var sheetName, out var rangeAddress))
        {
            return -1;
        }

        return WriteResult(_formattingCommands.ClearRules(batch, sheetName, rangeAddress));
    }

    private bool TryGetRangeInputs(Settings settings, out string sheetName, out string rangeAddress)
    {
        sheetName = settings.SheetName?.Trim() ?? string.Empty;
        rangeAddress = settings.RangeAddress?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            _console.WriteError("--range is required for this action.");
            return false;
        }

        return true;
    }

    private static bool RequiresSecondFormula(string? operatorType)
    {
        if (string.IsNullOrWhiteSpace(operatorType))
        {
            return false;
        }

        return operatorType.Equals("between", StringComparison.OrdinalIgnoreCase) ||
               operatorType.Equals("notbetween", StringComparison.OrdinalIgnoreCase) ||
               operatorType.Equals("not-between", StringComparison.OrdinalIgnoreCase);
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown conditionalformatting action '{action}'.");
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

        [CommandOption("--range <RANGE>")]
        public string? RangeAddress { get; init; }

        [CommandOption("--rule-type <TYPE>")]
        public string? RuleType { get; init; }

        [CommandOption("--operator <OPERATOR>")]
        public string? OperatorType { get; init; }

        [CommandOption("--formula1 <FORMULA>")]
        public string? Formula1 { get; init; }

        [CommandOption("--formula2 <FORMULA>")]
        public string? Formula2 { get; init; }

        [CommandOption("--interior-color <COLOR>")]
        public string? InteriorColor { get; init; }

        [CommandOption("--interior-pattern <PATTERN>")]
        public string? InteriorPattern { get; init; }

        [CommandOption("--font-color <COLOR>")]
        public string? FontColor { get; init; }

        [CommandOption("--font-bold <BOOL>")]
        public bool? FontBold { get; init; }

        [CommandOption("--font-italic <BOOL>")]
        public bool? FontItalic { get; init; }

        [CommandOption("--border-style <STYLE>")]
        public string? BorderStyle { get; init; }

        [CommandOption("--border-color <COLOR>")]
        public string? BorderColor { get; init; }
    }
}
