using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Text;
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
        output.AppendLine(CultureInfo.InvariantCulture, $"Usage: excelcli {arguments[0]} {arguments[1]} [OPTIONS]");
        output.AppendLine();
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
        }
        output.AppendLine(CultureInfo.InvariantCulture, $"  {"-h|--help",-42} Show this action-specific help.");

        CliCommandRuntime.Current.Output.Write(output.ToString());
        return true;
    }
}
