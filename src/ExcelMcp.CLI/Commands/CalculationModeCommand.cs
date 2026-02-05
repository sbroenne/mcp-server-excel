using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Calculation mode commands - thin wrapper that sends requests to service.
/// Actions: get-mode, set-mode, calculate
/// </summary>
internal sealed class CalculationModeCommand : ServiceCommandBase<ServiceRegistry.Calculation.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.Calculation.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.Calculation.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.Calculation.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.Calculation.CliSettings settings, string action)
    {
        return ServiceRegistry.Calculation.RouteCliArgs(
            action,
            mode: settings.Mode,
            scope: settings.Scope,
            sheetName: settings.SheetName,
            rangeAddress: settings.RangeAddress);
    }
}


