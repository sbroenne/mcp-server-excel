using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Connection commands - thin wrapper that sends requests to service.
/// Actions: list, view, create, test, refresh, delete, load-to, get-properties, set-properties
/// </summary>
internal sealed class ConnectionCommand : ServiceCommandBase<ServiceRegistry.Connection.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.Connection.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.Connection.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.Connection.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.Connection.CliSettings settings, string action)
    {
        return ServiceRegistry.Connection.RouteCliArgs(
            action,
            connectionName: settings.ConnectionName,
            connectionString: settings.ConnectionString,
            commandText: settings.CommandText,
            description: settings.Description,
            timeout: settings.Timeout,
            sheetName: settings.SheetName,
            backgroundQuery: settings.BackgroundQuery,
            refreshOnFileOpen: settings.RefreshOnFileOpen,
            savePassword: settings.SavePassword,
            refreshPeriod: settings.RefreshPeriod
        );
    }
}


