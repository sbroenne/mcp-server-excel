using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// DataModel commands - uses generated CliSettings and RouteCliArgs.
/// </summary>
internal sealed class DataModelCommand : ServiceCommandBase<ServiceRegistry.DataModel.CliSettings>
{
    protected override string? GetSessionId(ServiceRegistry.DataModel.CliSettings settings) => settings.SessionId;
    protected override string? GetAction(ServiceRegistry.DataModel.CliSettings settings) => settings.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.DataModel.ValidActions;

    protected override (string command, object? args) Route(ServiceRegistry.DataModel.CliSettings settings, string action)
    {
        // Resolve DAX/DMV from file if provided
        var daxFormula = ParameterTransforms.ResolveFileOrValue(settings.DaxFormula, settings.DaxFormulaFile);
        var daxQuery = ParameterTransforms.ResolveFileOrValue(settings.DaxQuery, settings.DaxQueryFile);
        var dmvQuery = ParameterTransforms.ResolveFileOrValue(settings.DmvQuery, settings.DmvQueryFile);

        return ServiceRegistry.DataModel.RouteCliArgs(
            action,
            tableName: settings.TableName,
            measureName: settings.MeasureName,
            oldName: settings.OldName,
            newName: settings.NewName,
            timeout: settings.Timeout,
            daxFormula: daxFormula,
            formatType: settings.FormatType,
            description: settings.Description,
            daxQuery: daxQuery,
            dmvQuery: dmvQuery
        );
    }
}


