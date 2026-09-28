using System.Reflection;
using Xunit;

namespace Sbroenne.ExcelMcp.Tests.Infrastructure;

[Trait("Category", "Architecture")]
[Trait("RequiresExcel", "false")]
public sealed class TestClassificationArchitectureTests
{
    [Fact]
    public void DiscoveredTestsHaveOneExcelClassificationAndExcelTestsAreExclusive()
    {
        var assembly = typeof(TestClassificationArchitectureTests).Assembly;
        var exclusiveCollections = assembly.GetTypes()
            .SelectMany(type => type.CustomAttributes)
            .Where(attribute => attribute.AttributeType == typeof(CollectionDefinitionAttribute)
                && attribute.NamedArguments.Any(argument =>
                    argument.MemberName == nameof(CollectionDefinitionAttribute.DisableParallelization)
                    && argument.TypedValue.Value is true))
            .Select(attribute => (string)attribute.ConstructorArguments[0].Value!)
            .ToHashSet(StringComparer.Ordinal);
        var errors = new List<string>();

        foreach (var type in assembly.GetTypes().Where(type => type.IsClass))
        {
            foreach (var method in type.GetMethods(
                BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                .Where(IsTestMethod))
            {
                var classifications = type.CustomAttributes
                    .Concat(method.CustomAttributes)
                    .Where(attribute => attribute.AttributeType == typeof(TraitAttribute)
                        && string.Equals(
                            (string?)attribute.ConstructorArguments[0].Value,
                            "RequiresExcel",
                            StringComparison.OrdinalIgnoreCase))
                    .Select(attribute =>
                        ((string)attribute.ConstructorArguments[1].Value!).ToLowerInvariant())
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();
                var testName = $"{type.FullName}.{method.Name}";

                if (classifications.Length != 1
                    || classifications[0] is not ("true" or "false"))
                {
                    errors.Add(
                        $"{testName}: expected exactly one RequiresExcel=true/false classification.");
                    continue;
                }

                if (classifications[0] == "true")
                {
                    var collection = type.CustomAttributes
                        .Where(attribute => attribute.AttributeType == typeof(CollectionAttribute))
                        .Select(attribute => (string?)attribute.ConstructorArguments[0].Value)
                        .SingleOrDefault();
                    if (collection is null || !exclusiveCollections.Contains(collection))
                    {
                        errors.Add(
                            $"{testName}: Excel tests must use a collection definition with DisableParallelization=true.");
                    }
                }
            }
        }

        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));
    }

    private static bool IsTestMethod(MethodInfo method) =>
        method.GetCustomAttributes(inherit: true).Any(attribute => attribute is FactAttribute);
}
