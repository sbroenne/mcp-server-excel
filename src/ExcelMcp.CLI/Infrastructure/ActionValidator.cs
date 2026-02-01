using System.Collections.Concurrent;
using System.Reflection;
using Sbroenne.ExcelMcp.Core.Models.Actions;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class ActionValidator
{
    private static readonly ConcurrentDictionary<Type, IReadOnlyCollection<string>> ActionCache = new();

    public static IReadOnlyCollection<string> GetValidActions<TEnum>() where TEnum : struct, Enum
    {
        return ActionCache.GetOrAdd(typeof(TEnum), BuildActionsForEnum);
    }

    public static bool TryNormalizeAction<TEnum>(string actionInput, out string normalizedAction, out string errorMessage)
        where TEnum : struct, Enum
    {
        return TryNormalizeAction(actionInput, GetValidActions<TEnum>(), out normalizedAction, out errorMessage);
    }

    public static bool TryNormalizeAction(string actionInput, IEnumerable<string> validActions, out string normalizedAction, out string errorMessage)
    {
        normalizedAction = actionInput.Trim().ToLowerInvariant();
        var validSet = validActions.ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (validSet.Contains(normalizedAction))
        {
            errorMessage = string.Empty;
            return true;
        }

        errorMessage = $"Invalid action '{normalizedAction}'. Valid actions: {string.Join(", ", validSet.OrderBy(a => a, StringComparer.OrdinalIgnoreCase))}";
        return false;
    }

    private static IReadOnlyCollection<string> BuildActionsForEnum(Type enumType)
    {
        var toActionString = typeof(ActionExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .FirstOrDefault(m => m.Name == "ToActionString" && m.GetParameters().Length == 1 && m.GetParameters()[0].ParameterType == enumType);

        if (toActionString == null)
        {
            throw new InvalidOperationException($"Missing ToActionString() for enum {enumType.Name}");
        }

        var values = Enum.GetValues(enumType);
        var results = new List<string>(values.Length);

        foreach (var value in values)
        {
            var result = toActionString.Invoke(null, [value]) as string;
            if (string.IsNullOrWhiteSpace(result))
            {
                throw new InvalidOperationException($"Action mapping missing for {enumType.Name}.{value}");
            }
            results.Add(result);
        }

        return results;
    }
}
