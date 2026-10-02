using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacRangeArguments
{
    public static void Prepare(string category, string action, JsonObject arguments)
    {
        if (category != "range")
        {
            return;
        }

        if (action == "set-values")
        {
            var values = arguments["values"]?.Deserialize<List<List<object?>>>(ServiceProtocol.JsonOptions);
            var valuesFile = arguments["valuesFile"]?.GetValue<string>();
            arguments["values"] = JsonSerializer.SerializeToNode(
                ParameterTransforms.ResolveValuesOrFile(values, valuesFile),
                ServiceProtocol.JsonOptions);
            arguments.Remove("valuesFile");
        }
        else if (action == "set-formulas")
        {
            var formulas = arguments["formulas"]?.Deserialize<List<List<string>>>(ServiceProtocol.JsonOptions);
            var formulasFile = arguments["formulasFile"]?.GetValue<string>();
            arguments["formulas"] = JsonSerializer.SerializeToNode(
                ParameterTransforms.ResolveFormulasOrFile(formulas, formulasFile),
                ServiceProtocol.JsonOptions);
            arguments.Remove("formulasFile");
        }
        else if (action == "set-number-formats")
        {
            var formats = arguments["formats"]?.Deserialize<List<List<string>>>(ServiceProtocol.JsonOptions);
            var formatsFile = arguments["formatsFile"]?.GetValue<string>();
            arguments["formats"] = JsonSerializer.SerializeToNode(
                ParameterTransforms.ResolveFormulasOrFile(formats, formatsFile, "formats"),
                ServiceProtocol.JsonOptions);
            arguments.Remove("formatsFile");
        }
    }
}
