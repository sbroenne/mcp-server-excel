using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// PowerQuery commands - uses generated CliSettings and RouteCliArgs.
/// </summary>
internal sealed class PowerQueryCommand : ServiceCommandBase<ServiceRegistry.PowerQuery.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.PowerQuery.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.PowerQuery.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.PowerQuery.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.PowerQuery.CliSettings settings, string action)
    {
        // Resolve M code from file if provided
        var mCode = ParameterTransforms.ResolveFileOrValue(settings.MCode, settings.MCodeFile);

        return ServiceRegistry.PowerQuery.RouteCliArgs(
            action,
            queryName: settings.QueryName,
            mCode: mCode,
            loadDestination: settings.LoadDestination,
            targetSheet: settings.TargetSheet,
            targetCellAddress: settings.TargetCellAddress,
            oldName: settings.OldName,
            newName: settings.NewName
        );
    }
}


