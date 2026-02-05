using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// ConditionalFormat commands - thin wrapper that sends requests to service.
/// Actions: add-rule, clear-rules
/// </summary>
internal sealed class ConditionalFormatCommand : ServiceCommandBase<ServiceRegistry.ConditionalFormat.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.ConditionalFormat.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.ConditionalFormat.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.ConditionalFormat.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.ConditionalFormat.CliSettings settings, string action)
    {
        return ServiceRegistry.ConditionalFormat.RouteCliArgs(
            action,
            sheetName: settings.SheetName,
            rangeAddress: settings.RangeAddress,
            ruleType: settings.RuleType,
            operatorType: settings.OperatorType,
            formula1: settings.Formula1,
            formula2: settings.Formula2,
            interiorColor: settings.InteriorColor,
            interiorPattern: settings.InteriorPattern,
            fontColor: settings.FontColor,
            fontBold: settings.FontBold,
            fontItalic: settings.FontItalic,
            borderStyle: settings.BorderStyle,
            borderColor: settings.BorderColor
        );
    }
}


