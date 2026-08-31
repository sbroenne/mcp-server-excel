using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.Generated;
using Spectre.Console.Cli;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ActionValidation")]
[Trait("Speed", "Fast")]
[Collection("ConsoleOutput")]
public sealed class ActionValidatorTests
{
    public static IEnumerable<object[]> ActionEnumTypes =>
    [
        [typeof(RangeAction), typeof(ServiceRegistry.Range)],
        [typeof(RangeEditAction), typeof(ServiceRegistry.RangeEdit)],
        [typeof(RangeFormatAction), typeof(ServiceRegistry.RangeFormat)],
        [typeof(RangeLinkAction), typeof(ServiceRegistry.RangeLink)],
        [typeof(DrawingAction), typeof(ServiceRegistry.Drawing)],
        [typeof(TableAction), typeof(ServiceRegistry.Table)],
        [typeof(WorkbookAction), typeof(ServiceRegistry.Workbook)]
    ];

    private static readonly string[] ExpectedCommands =
    [
        "session",
        "workbook",
        "sheet",
        "worksheetstyle",
        "range",
        "rangeedit",
        "rangeformat",
        "rangelink",
        "table",
        "tablecolumn",
        "powerquery",
        "pivottable",
        "pivottablefield",
        "pivottablecalc",
        "chart",
        "chartconfig",
        "connection",
        "calculationmode",
        "namedrange",
        "conditionalformat",
        "vba",
        "datamodel",
        "datamodelrelationship",
        "drawing",
        "slicer"
    ];

    [Theory]
    [MemberData(nameof(ActionEnumTypes))]
    public void GetValidActions_ReturnsAllActionStrings(Type enumType, Type registryType)
    {
        var expected = GetExpectedActions(enumType, registryType);
        var actual = GetActualActions(registryType);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ListActionsCommand_AllCommands_ReturnsExpectedKeys()
    {
        var command = new ListActionsCommand();
        var settings = new ListActionsCommand.Settings();

        var context = new CommandContext(
            Array.Empty<string>(),
            new FakeRemainingArguments(),
            "actions",
            null);
        var executeMethod = typeof(ListActionsCommand).GetMethod(
            "Execute",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        var output = CaptureOutput(() => (int)executeMethod.Invoke(command, [context, settings, CancellationToken.None])!);
        using var document = JsonDocument.Parse(output);

        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        var commands = document.RootElement.GetProperty("commands");

        foreach (var expected in ExpectedCommands)
        {
            Assert.True(commands.TryGetProperty(expected, out _), $"Missing command '{expected}'.");
        }
    }

    [Fact]
    public void DataModelActions_ReturnAllGeneratedActionStrings()
    {
        var expected = GetExpectedActions(typeof(DataModelAction), typeof(ServiceRegistry.DataModel));
        var actual = GetActualActions(typeof(ServiceRegistry.DataModel));

        Assert.Equal(expected, actual);
        Assert.Contains("read-connection", actual);
    }

    [Fact]
    public void TableActions_IncludePreflight()
    {
        var actions = GetActualActions(typeof(ServiceRegistry.Table));

        Assert.Contains("preflight", actions);
    }

    [Fact]
    public void SheetDescription_DoesNotReferenceUnregisteredStyleCommand()
    {
        Assert.DoesNotContain(
            "Use sheetstyle",
            ServiceRegistry.Sheet.Description,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GeneratedCommandDescriptions_AreConcise()
    {
        var offenders = typeof(ServiceRegistry)
            .GetNestedTypes(BindingFlags.Public)
            .Select(type => new
            {
                type.Name,
                Description = type
                    .GetField("Description", BindingFlags.Public | BindingFlags.Static)?
                    .GetRawConstantValue() as string
            })
            .Where(item => item.Description is { Length: > 180 })
            .Select(item => $"{item.Name} ({item.Description!.Length} characters)")
            .ToArray();

        Assert.True(
            offenders.Length == 0,
            "Top-level CLI command descriptions must stay under 180 characters: " +
            string.Join(", ", offenders));
    }

    private static string[] GetExpectedActions(Type enumType, Type registryType)
    {
        // Find ToActionString method in the ServiceRegistry nested type (e.g., ServiceRegistry.Range.ToActionString)
        var actionMethod = registryType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .First(m => m.Name == "ToActionString" && m.GetParameters().Length == 1 && m.GetParameters()[0].ParameterType == enumType);

        var values = Enum.GetValues(enumType);
        var results = new List<string>(values.Length);

        foreach (var value in values)
        {
            var action = actionMethod.Invoke(null, [value]) as string;
            results.Add(action ?? string.Empty);
        }

        return results.OrderBy(action => action, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static string[] GetActualActions(Type registryType)
    {
        // Get ValidActions field from the ServiceRegistry nested type (e.g., ServiceRegistry.Range.ValidActions)
        var validActionsField = registryType
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .First(f => f.Name == "ValidActions");

        var actions = (string[])validActionsField.GetValue(null)!;
        return actions.OrderBy(action => action, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static string CaptureOutput(Func<int> action)
    {
        var original = Console.Out;
        using var writer = new StringWriter();
        try
        {
            Console.SetOut(writer);
            action();
            return writer.ToString().Trim();
        }
        finally
        {
            Console.SetOut(original);
        }
    }

    private sealed class FakeRemainingArguments : IRemainingArguments
    {
        private static readonly ILookup<string, string?> EmptyLookup =
            Array.Empty<string>().ToLookup(_ => string.Empty, _ => (string?)null);

        public ILookup<string, string?> Parsed { get; } = EmptyLookup;
        public IReadOnlyList<string> Raw { get; } = Array.Empty<string>();
    }
}
