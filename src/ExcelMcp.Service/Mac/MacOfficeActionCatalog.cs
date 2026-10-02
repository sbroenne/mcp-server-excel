using System.Reflection;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacOfficeAction(
    string Command,
    string RequirementSet,
    bool Mutation);

internal static class MacOfficeActionCatalog
{
    private static readonly Dictionary<string, MacOfficeAction> Actions = Build();

    public static bool TryGet(string command, out MacOfficeAction action) =>
        Actions.TryGetValue(command, out action!);

    public static IReadOnlyCollection<MacOfficeAction> All => Actions.Values;

    private static Dictionary<string, MacOfficeAction> Build()
    {
        var actions = new Dictionary<string, MacOfficeAction>(StringComparer.Ordinal);
        foreach (var contract in typeof(ServiceCategoryAttribute).Assembly.GetTypes()
                     .Where(type => type.IsInterface))
        {
            var category = contract.GetCustomAttribute<ServiceCategoryAttribute>();
            if (category is null)
            {
                continue;
            }

            foreach (var method in contract.GetMethods())
            {
                var office = method.GetCustomAttribute<OfficeAddInActionAttribute>();
                if (office is null)
                {
                    continue;
                }

                var service = method.GetCustomAttribute<ServiceActionAttribute>()
                    ?? throw new InvalidOperationException(
                        $"{contract.Name}.{method.Name} has OfficeAddInAction without ServiceAction.");
                var command = $"{category.Category}.{service.Action}";
                actions.Add(command, new MacOfficeAction(
                    command,
                    office.RequirementSet,
                    office.Mutation));
            }
        }

        return actions;
    }
}
