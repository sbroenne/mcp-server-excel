using System.Reflection;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Tests.Infrastructure;

[Trait("Category", "Architecture")]
[Trait("RequiresExcel", "false")]
public sealed class TestClassificationArchitectureTests
{
    [Fact]
    public void ExcelValidationGroups_PreserveClassFixturesAndExportBuiltInventory()
    {
        var assembly = typeof(TestClassificationArchitectureTests).Assembly;
        var selections = new List<ExcelSelection>();
        foreach (var type in assembly.GetTypes().Where(type => type.IsClass))
        {
            var methods = type.GetMethods().Where(IsTestMethod).ToArray();
            var classFeatures = methods.SelectMany(method => Traits(type, method, "Feature"))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var classGroup = GetExcelGroup(type, classFeatures);
            var classSelections = methods
                .Where(method => Traits(type, method, "RequiresExcel").Contains("true", StringComparer.OrdinalIgnoreCase))
                .Select(method => new ExcelSelection(
                    type.FullName!,
                    $"{type.FullName}.{method.Name}",
                    method.Name.StartsWith("VbaRun_OnMacroWorkbook", StringComparison.Ordinal) ? "VBA" : classGroup,
                    Traits(type, method, "RunType").Contains("OnDemand", StringComparer.OrdinalIgnoreCase),
                    method.GetCustomAttributes(inherit: true).Any(attribute =>
                        attribute.GetType().Name is "ConfiguredIrmFactAttribute" or "JapaneseLocaleFactAttribute"),
                    Traits(type, method, "Acceptance").Contains("Required", StringComparer.Ordinal)))
                .ToArray();
            if (type.GetInterfaces().Any(item => item.IsGenericType &&
                item.GetGenericTypeDefinition() == typeof(IClassFixture<>)))
            {
                Assert.True(classSelections.Select(item => item.Group).Distinct(StringComparer.Ordinal).Count() <= 1,
                    $"{type.FullName}: splitting a class fixture across Excel groups would repeat its setup.");
            }
            selections.AddRange(classSelections);
        }
        var output = Environment.GetEnvironmentVariable("EXCELMCP_TEST_SELECTION_OUTPUT");
        if (!string.IsNullOrEmpty(output))
        {
            File.WriteAllText(output, JsonSerializer.Serialize(selections));
        }
    }

    [Theory]
    [InlineData("Range,Ranges,Tables", "Editing")]
    [InlineData("Tables,DataModel", "Data")]
    [InlineData("PowerQuery,DataModel,PivotTables", "Data")]
    [InlineData("Charts,PivotTables", "Reporting")]
    [InlineData("Window,Screenshot", "Desktop")]
    [InlineData("VBA", "VBA")]
    [InlineData("SessionLifecycle", "Lifecycle")]
    public void ExcelGroups_UseAllFeatureSpellingsAndCrossFeaturePrecedence(string features, string expected)
    {
        Assert.Equal(expected, GetExcelGroup(typeof(object),
            features.Split(',').ToHashSet(StringComparer.OrdinalIgnoreCase)));
    }

    private sealed record ExcelSelection(
        string Class, string Method, string Group, bool OnDemand, bool Prerequisite, bool Required);

    private static IEnumerable<string> Traits(Type type, MethodInfo method, string name) =>
        type.CustomAttributes.Concat(method.CustomAttributes)
            .Where(attribute => attribute.AttributeType == typeof(TraitAttribute) &&
                string.Equals((string?)attribute.ConstructorArguments[0].Value, name, StringComparison.OrdinalIgnoreCase))
            .Select(attribute => (string)attribute.ConstructorArguments[1].Value!);

    private static string GetExcelGroup(Type type, HashSet<string> features)
    {
        if (type.Assembly.GetName().Name == "Sbroenne.ExcelMcp.ComInterop.Tests") { return "Infrastructure"; }
        if (type.Name is "PreBuildGracefulSaveAcceptanceTests" or "CliWorkflowAcceptanceTests" or "McpServerSmokeTests" ||
            type.CustomAttributes.Any(attribute => attribute.AttributeType == typeof(TraitAttribute) &&
                (string?)attribute.ConstructorArguments[0].Value == "Acceptance" &&
                (string?)attribute.ConstructorArguments[1].Value == "Required")) { return "Acceptance"; }
        if (features.Overlaps(["VBA", "VBATrust"])) { return "VBA"; }
        if (features.Overlaps(["Screenshot", "Window"])) { return "Desktop"; }
        if (features.Overlaps(["PowerQuery", "Connection", "Connections", "QueryTable", "DataModel", "PythonInExcel"])) { return "Data"; }
        if (features.Overlaps(["Chart", "Charts", "PivotTable", "PivotTables", "Slicer", "Drawing", "Analysis", "CalculationMode", "ConditionalFormat"])) { return "Reporting"; }
        if (features.Overlaps(["Range", "Ranges", "Sheet", "Worksheets", "Table", "Tables", "Parameters", "XmlMap", "Formatting"])) { return "Editing"; }
        return "Lifecycle";
    }

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
                if (Traits(type, method, "AdapterTestKind").Contains("System", StringComparer.Ordinal) &&
                    assembly.GetName().Name != "Sbroenne.ExcelMcp.CLI.Tests")
                {
                    errors.Add($"{testName}: add this project's system tests to the hosted Process selection.");
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
