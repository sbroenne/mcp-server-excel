using System.Collections;
using System.Reflection;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Rejects values that Spectre.Console.Cli inserted itself. When strict parsing fails, Spectre retries by
/// inserting its internal default-command name at each position, so an option or argument whose value the
/// user omitted can silently receive that name instead of failing.
/// </summary>
internal sealed class MissingValueInterceptor(IReadOnlyList<string> userArguments) : ICommandInterceptor
{
    internal const string SpectreDefaultCommandName = "__default_command";

    public void Intercept(CommandContext context, CommandSettings settings)
    {
        var properties = settings.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var options = new Dictionary<string, (PropertyInfo Property, bool TakesValue)>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            if (property.GetCustomAttribute<CommandOptionAttribute>() is not { } option)
                continue;
            foreach (var name in option.LongNames.Select(name => "--" + name).Concat(option.ShortNames.Select(name => "-" + name)))
                options[name] = (property, option.ValueName is not null);
        }

        var (typedOptionValues, typedPositionalValues) = CountTypedPlaceholders(options);

        foreach (var property in properties)
        {
            var inserted = CountPlaceholders(property.GetValue(settings));
            if (inserted == 0)
                continue;
            if (property.GetCustomAttribute<CommandOptionAttribute>() is { } option)
            {
                if (inserted <= typedOptionValues.GetValueOrDefault(property))
                    continue;
                var names = option.LongNames.Select(name => "--" + name).Concat(option.ShortNames.Select(name => "-" + name)).ToArray();
                var typed = names.FirstOrDefault(name => userArguments.Contains(name, StringComparer.OrdinalIgnoreCase)) ?? names[0];
                throw new ArgumentException($"Option '{typed}' requires a value.");
            }
            if (property.GetCustomAttribute<CommandArgumentAttribute>() is { } argument && inserted > typedPositionalValues)
                throw new ArgumentException($"Argument '<{argument.ValueName}>' is required.");
        }
    }

    private (Dictionary<PropertyInfo, int> OptionValues, int PositionalValues) CountTypedPlaceholders(
        Dictionary<string, (PropertyInfo Property, bool TakesValue)> options)
    {
        var optionValues = new Dictionary<PropertyInfo, int>();
        var positionalValues = 0;
        for (var i = 0; i < userArguments.Count; i++)
        {
            var current = userArguments[i];
            // Spectre keeps words after "--" as remaining arguments, never as settings values.
            if (current == "--")
                break;
            if (!current.StartsWith('-'))
            {
                if (current == SpectreDefaultCommandName)
                    positionalValues++;
                continue;
            }

            var separator = current.IndexOfAny(['=', ':']);
            var name = separator < 0 ? current : current[..separator];
            if (!options.TryGetValue(name, out var option) || !option.TakesValue)
                continue;
            string? value = separator >= 0
                ? current[(separator + 1)..]
                : i + 1 < userArguments.Count && !options.ContainsKey(userArguments[i + 1]) && userArguments[i + 1] != "--"
                    ? userArguments[++i]
                    : null;
            if (value == SpectreDefaultCommandName)
                optionValues[option.Property] = optionValues.GetValueOrDefault(option.Property) + 1;
        }
        return (optionValues, positionalValues);
    }

    private static int CountPlaceholders(object? value) => value switch
    {
        string text => text == SpectreDefaultCommandName ? 1 : 0,
        IEnumerable values => values.OfType<string>().Count(item => item == SpectreDefaultCommandName),
        _ => 0
    };
}
