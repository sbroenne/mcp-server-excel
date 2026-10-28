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
        foreach (var property in settings.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!ContainsInsertedName(property.GetValue(settings)))
                continue;
            if (property.GetCustomAttribute<CommandOptionAttribute>() is { } option)
            {
                var names = option.LongNames.Select(name => "--" + name).Concat(option.ShortNames.Select(name => "-" + name)).ToArray();
                if (TypedAsOptionValue(names))
                    continue;
                var typed = names.FirstOrDefault(name => userArguments.Contains(name, StringComparer.OrdinalIgnoreCase)) ?? names[0];
                throw new ArgumentException($"Option '{typed}' requires a value.");
            }
            if (property.GetCustomAttribute<CommandArgumentAttribute>() is { } argument)
            {
                if (TypedAsArgument())
                    continue;
                throw new ArgumentException($"Argument '<{argument.ValueName}>' is required.");
            }
        }
    }

    private bool TypedAsOptionValue(string[] names)
    {
        for (var i = 0; i < userArguments.Count; i++)
        {
            var current = userArguments[i];
            if (names.Contains(current, StringComparer.OrdinalIgnoreCase)
                && i + 1 < userArguments.Count
                && userArguments[i + 1] == SpectreDefaultCommandName)
                return true;
            if (names.Any(name => current.Equals(name + "=" + SpectreDefaultCommandName, StringComparison.OrdinalIgnoreCase)
                || current.Equals(name + ":" + SpectreDefaultCommandName, StringComparison.OrdinalIgnoreCase)))
                return true;
        }
        return false;
    }

    private bool TypedAsArgument() =>
        Enumerable.Range(0, userArguments.Count).Any(i =>
            userArguments[i] == SpectreDefaultCommandName && (i == 0 || !userArguments[i - 1].StartsWith('-')));

    private static bool ContainsInsertedName(object? value) => value switch
    {
        string text => text == SpectreDefaultCommandName,
        IEnumerable values => values.OfType<string>().Contains(SpectreDefaultCommandName, StringComparer.Ordinal),
        _ => false
    };
}
