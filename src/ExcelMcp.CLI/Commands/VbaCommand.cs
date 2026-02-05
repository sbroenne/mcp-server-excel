using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// VBA commands - uses generated CliSettings and RouteCliArgs.
/// </summary>
internal sealed class VbaCommand : ServiceCommandBase<ServiceRegistry.Vba.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.Vba.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.Vba.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.Vba.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.Vba.CliSettings settings, string action)
    {
        // Resolve VBA code from file if provided
        var vbaCode = ParameterTransforms.ResolveFileOrValue(settings.VbaCode, settings.VbaCodeFile);

        return ServiceRegistry.Vba.RouteCliArgs(
            action,
            moduleName: settings.ModuleName,
            vbaCode: vbaCode,
            procedureName: settings.ProcedureName,
            timeout: settings.Timeout,
            parameters: settings.Parameters);
    }
}


