using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class ActionHelp
{
    internal static bool TryWrite(string[] arguments)
    {
        if (arguments.Length < 3
            || !arguments.Any(argument => argument is "--help" or "-h")
            || !_CliCategoryMetadata.ActionParametersByCommand.TryGetValue(arguments[0], out var actions)
            || !actions.TryGetValue(arguments[1], out var actionParameters))
        {
            return false;
        }

        var category = _CliCategoryMetadata.Categories.Single(
            item => string.Equals(item.CliCommandName, arguments[0], StringComparison.OrdinalIgnoreCase));
        var registryTypeName = category.RegistryTypeName.Split('.').Last();
        var registryType = typeof(ServiceRegistry).GetNestedType(registryTypeName, BindingFlags.Public)
            ?? throw new InvalidOperationException($"Generated service registry type '{registryTypeName}' was not found.");
        var settingsType = registryType.GetNestedType("CliSettings", BindingFlags.Public)
            ?? throw new InvalidOperationException($"Generated CLI settings for '{arguments[0]}' were not found.");

        var parameterNames = new List<string>(actionParameters);
        if (category.RequiresSession)
            parameterNames.Insert(0, "SessionId");
        parameterNames.Add("OutputPath");

        var output = new StringBuilder();
        var batchMappings = new List<(string Option, string Parameter)>();
        output.AppendLine(CultureInfo.InvariantCulture, $"Usage: excelcli {arguments[0]} {arguments[1]} [OPTIONS]");
        output.AppendLine();
        var actionDescription = _CliCategoryMetadata.ActionDescriptionsByCommand[arguments[0]][arguments[1]];
        if (!string.IsNullOrWhiteSpace(actionDescription))
        {
            output.AppendLine(actionDescription);
            output.AppendLine();
        }
        output.AppendLine("Options:");
        foreach (var parameterName in parameterNames)
        {
            var property = settingsType.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .SingleOrDefault(candidate => string.Equals(
                    candidate.Name, parameterName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException(
                    $"Generated CLI option '{parameterName}' was not found for '{arguments[0]}'.");
            var optionAttribute = property.GetCustomAttributesData().SingleOrDefault(attribute =>
                attribute.AttributeType.FullName == "Spectre.Console.Cli.CommandOptionAttribute")
                ?? throw new InvalidOperationException(
                    $"Generated CLI option '{parameterName}' has no command-line definition.");
            var option = optionAttribute.ConstructorArguments[0].Value?.ToString()
                ?? throw new InvalidOperationException(
                    $"Generated CLI option '{parameterName}' has no command-line name.");
            var description = property.GetCustomAttribute<DescriptionAttribute>()?.Description ?? string.Empty;
            description = description
                .Replace("[[", "[", StringComparison.Ordinal)
                .Replace("]]", "]", StringComparison.Ordinal);
            output.AppendLine(CultureInfo.InvariantCulture, $"  {option,-42} {description}");
            if (actionParameters.Contains(parameterName, StringComparer.Ordinal))
                batchMappings.Add((option, parameterName));
        }
        output.AppendLine(CultureInfo.InvariantCulture, $"  {"-h|--help",-42} Show this action-specific help.");
        var actionName = actions.Keys.Single(action =>
            string.Equals(action, arguments[1], StringComparison.OrdinalIgnoreCase));
        var batchCommand = $"{category.CliCommandName}.{actionName}";
        output.AppendLine();
        output.AppendLine("Batch input (fill in required args from the options above):");
        output.AppendLine(JsonSerializer.Serialize(new
        {
            command = batchCommand,
            args = new Dictionary<string, object?>()
        }));
        output.AppendLine("Batch arguments (case-sensitive):");
        foreach (var (option, parameter) in batchMappings)
            output.AppendLine(CultureInfo.InvariantCulture, $"  {option,-42} args.{parameter}");
        if (batchMappings.Count == 0)
            output.AppendLine("  This action takes no args.");
        if (category.RequiresSession)
            output.AppendLine("Select the workbook with batch --session <id> or an entry's top-level sessionId, not inside args.");
        output.AppendLine("Use JSON arrays/objects directly, not strings containing JSON. See excelcli batch --help for execution examples.");
        if (batchCommand == "range.set-values")
        {
            output.AppendLine("Example args (requires an existing Sales sheet and authorization to write inspected empty A1:B2 cells):");
            output.AppendLine(BatchHelp.RangeWriteEntry);
        }

        CliCommandRuntime.Current.Output.Write(output.ToString());
        return true;
    }
}
