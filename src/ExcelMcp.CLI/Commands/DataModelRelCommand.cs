using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// DataModelRel commands - uses generated CliSettings and RouteCliArgs.
/// </summary>
internal sealed class DataModelRelCommand : ServiceCommandBase<ServiceRegistry.DataModelRel.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.DataModelRel.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.DataModelRel.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.DataModelRel.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.DataModelRel.CliSettings settings, string action)
    {
        return ServiceRegistry.DataModelRel.RouteCliArgs(
            action,
            fromTable: settings.FromTable,
            fromColumn: settings.FromColumn,
            toTable: settings.ToTable,
            toColumn: settings.ToColumn,
            active: settings.Active
        );
    }
}


