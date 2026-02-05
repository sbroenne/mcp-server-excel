using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// NamedRange commands - thin wrapper that sends requests to service.
/// Actions: list, read, write, create, update, delete
/// </summary>
internal sealed class NamedRangeCommand : ServiceCommandBase<ServiceRegistry.NamedRange.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.NamedRange.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.NamedRange.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.NamedRange.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.NamedRange.CliSettings settings, string action)
    {
        return ServiceRegistry.NamedRange.RouteCliArgs(
            action,
            paramName: settings.ParamName,
            value: settings.Value,
            reference: settings.Reference
        );
    }
}


