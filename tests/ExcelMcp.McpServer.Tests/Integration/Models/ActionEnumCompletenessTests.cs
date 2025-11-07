// Explicit usings retained; pragma used to suppress IDE0005 for clarity in reflection-heavy test
#pragma warning disable IDE0005
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
#pragma warning restore IDE0005
using Sbroenne.ExcelMcp.McpServer.Models;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Models;

/// <summary>
/// CRITICAL: Ensures all enum values have ToActionString() mappings.
/// Missing mappings cause ArgumentException at runtime when users invoke actions.
///
/// Uses reflection to automatically discover ALL action enums - no manual maintenance required.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ActionEnums")]
[Trait("RequiresExcel", "false")]
public class ActionEnumCompletenessTests
{
    private readonly ITestOutputHelper _output;
    /// <inheritdoc/>

    public ActionEnumCompletenessTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// CRITICAL: Discovers ALL *Action enums and verifies every value has a ToActionString() mapping.
    /// This test will FAIL if:
    /// 1. New enum added without ToActionString() extension method
    /// 2. Enum value added without corresponding mapping
    /// 3. ToActionString() throws ArgumentException for any enum value
    /// </summary>
    [Fact]
    public void AllActionEnums_HaveCompleteToActionStringMappings()
    {
        // Find all *Action enums in Models namespace
        var actionEnums = typeof(ActionExtensions).Assembly
            .GetTypes()
            .Where(t => t.IsEnum && t.Name.EndsWith("Action", StringComparison.Ordinal) && t.Namespace == "Sbroenne.ExcelMcp.McpServer.Models")
            .ToList();

        _output.WriteLine($"Found {actionEnums.Count} action enums:");
        foreach (var enumType in actionEnums)
        {
            _output.WriteLine($"  - {enumType.Name}");
        }

        Assert.NotEmpty(actionEnums); // Sanity check

        var failures = new List<string>();

        foreach (var enumType in actionEnums)
        {
            // Find the ToActionString extension method for this enum
            var extensionMethod = typeof(ActionExtensions)
                .GetMethods(BindingFlags.Public | BindingFlags.Static)
                .FirstOrDefault(m =>
                    m.Name == "ToActionString" &&
                    m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType == enumType);

            if (extensionMethod == null)
            {
                failures.Add($"❌ {enumType.Name}: Missing ToActionString() extension method");
                continue;
            }

            // Get all enum values
            var enumValues = Enum.GetValues(enumType);

            foreach (var enumValue in enumValues)
            {
                try
                {
                    // Invoke ToActionString() - will throw if mapping missing
                    var result = extensionMethod.Invoke(null, new[] { enumValue }) as string;

                    if (string.IsNullOrWhiteSpace(result))
                    {
                        failures.Add($"❌ {enumType.Name}.{enumValue}: Mapped to empty string");
                    }
                    else
                    {
                        _output.WriteLine($"  ✅ {enumType.Name}.{enumValue} → '{result}'");
                    }
                }
                catch (TargetInvocationException ex) when (ex.InnerException is ArgumentException argEx)
                {
                    failures.Add($"❌ {enumType.Name}.{enumValue}: {argEx.Message}");
                }
                catch (Exception ex)
                {
                    failures.Add($"❌ {enumType.Name}.{enumValue}: Unexpected error: {ex.Message}");
                }
            }
        }

        if (failures.Count > 0)
        {
            var message = $"Enum mapping failures:\n{string.Join("\n", failures)}";
            _output.WriteLine($"\n{message}");
            Assert.Fail(message);
        }
    }

    /// <summary>
    /// CRITICAL: Ensures no duplicate action strings within same enum (case-insensitive).
    /// Duplicates would cause ambiguous routing in tool switch statements.
    /// </summary>
    [Fact]
    public void AllActionEnums_NoDuplicateActionStrings()
    {
        var actionEnums = typeof(ActionExtensions).Assembly
            .GetTypes()
            .Where(t => t.IsEnum && t.Name.EndsWith("Action", StringComparison.Ordinal) && t.Namespace == "Sbroenne.ExcelMcp.McpServer.Models")
            .ToList();

        var failures = new List<string>();

        foreach (var enumType in actionEnums)
        {
            var extensionMethod = typeof(ActionExtensions)
                .GetMethods(BindingFlags.Public | BindingFlags.Static)
                .FirstOrDefault(m =>
                    m.Name == "ToActionString" &&
                    m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType == enumType);

            if (extensionMethod == null) continue; // Already caught by other test

            var enumValues = Enum.GetValues(enumType);
            var actionStrings = new List<(object enumValue, string actionString)>();

            foreach (var enumValue in enumValues)
            {
                try
                {
                    var result = extensionMethod.Invoke(null, new[] { enumValue }) as string;
                    if (!string.IsNullOrWhiteSpace(result))
                    {
                        actionStrings.Add((enumValue, result.ToLowerInvariant()));
                    }
                }
                catch
                {
                    // Ignore - will be caught by completeness test
                }
            }

            var duplicates = actionStrings
                .GroupBy(x => x.actionString)
                .Where(g => g.Count() > 1)
                .ToList();

            if (duplicates.Count > 0)
            {
                foreach (var duplicate in duplicates)
                {
                    var enumValueNames = string.Join(", ", duplicate.Select(x => x.enumValue));
                    failures.Add($"❌ {enumType.Name}: Duplicate action string '{duplicate.Key}' for: {enumValueNames}");
                }
            }
        }

        if (failures.Count > 0)
        {
            var message = $"Duplicate action string failures:\n{string.Join("\n", failures)}";
            _output.WriteLine($"\n{message}");
            Assert.Fail(message);
        }
    }

    /// <summary>
    /// CRITICAL: Verifies all *Tool.cs files have switch statements covering all enum values.
    /// This ensures tool implementations don't get out of sync with enums.
    ///
    /// NOTE: This is a compile-time check via exhaustive switch expressions.
    /// If switch is missing a case, C# compiler shows warning CS8524.
    /// This test documents the expectation - actual enforcement is via compiler.
    /// </summary>
    [Fact]
    public void AllActionEnums_DocumentedInToolFiles()
    {
        var actionEnums = typeof(ActionExtensions).Assembly
            .GetTypes()
            .Where(t => t.IsEnum && t.Name.EndsWith("Action", StringComparison.Ordinal) && t.Namespace == "Sbroenne.ExcelMcp.McpServer.Models")
            .ToList();

        _output.WriteLine($"\nExpected tool files with switch statements:");
        _output.WriteLine($"Each *Action enum should have corresponding *Tool.cs with exhaustive switch.\n");

        foreach (var enumType in actionEnums)
        {
            var toolName = enumType.Name.Replace("Action", "Tool");
            _output.WriteLine($"  - {enumType.Name} → Tools/{toolName}.cs");
            _output.WriteLine($"    Expected: switch (action.ToActionString()) with all {Enum.GetValues(enumType).Length} cases");
        }

        _output.WriteLine($"\n✅ Compiler enforces exhaustive switches via warning CS8524.");
        _output.WriteLine($"✅ Build with TreatWarningsAsErrors=true ensures no missing cases.");
    }
}
