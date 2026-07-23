using System.Text.Json;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for ServiceRegistry.DeserializeNestedCollection (generated helper).
/// Covers bug regressions for issue #521:
///   --values inline JSON loses string quotes in PowerShell (stdin sentinel + better error).
/// These tests must FAIL before the fix and PASS after.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceRegistry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceRegistryJsonParsingTests
{
    private static readonly System.Type _registryType = typeof(ServiceRegistry);

    private static List<List<object?>> Deserialize(string json)
    {
        var method = _registryType.GetMethod(
            "DeserializeNestedCollection",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var genericMethod = method.MakeGenericMethod(typeof(List<List<object?>>));
        try
        {
            return (List<List<object?>>)genericMethod.Invoke(null, [json])!;
        }
        catch (System.Reflection.TargetInvocationException ex)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException!).Throw();
            throw; // unreachable
        }
    }

    /// <summary>
    /// Regression test for issue #521 — error path:
    /// When PowerShell strips double-quotes, `--values '[["ACD Full Term",0.26]]'`
    /// arrives as `[[ACD Full Term,0.26]]` (invalid JSON).
    /// The error message MUST mention `--values-file` or `--values -` to help users.
    /// Before fix: error message says "Invalid JSON for nested collection. Expected 2D array..."
    ///             with no mention of PowerShell or --values-file.
    /// After fix: error message explicitly mentions the workaround options.
    /// </summary>
    [Fact]
    public void DeserializeNestedCollection_MangledJsonFromPowerShell_ErrorMentionsValuesFile()
    {
        // This is what PowerShell hands to the native exe after stripping inner quotes:
        var mangledJson = "[[ACD Full Term,0.26]]";

        var ex = Assert.Throws<ArgumentException>(() => Deserialize(mangledJson));

        // The error message must guide the user toward the correct fix
        Assert.True(
            ex.Message.Contains("values-file", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("--values -", StringComparison.OrdinalIgnoreCase),
            $"Error message should mention '--values-file' or '--values -' but was: {ex.Message}");
    }

    /// <summary>
    /// Regression test for issue #521 — stdin sentinel path:
    /// When `json == "-"`, the method should read JSON from Console.In instead of
    /// attempting to parse the literal string "-" as JSON.
    /// Before fix: throws ArgumentException because "-" is not valid JSON.
    /// After fix: reads from Console.In and deserializes correctly.
    /// </summary>
    [Fact]
    public void DeserializeNestedCollection_StdinSentinel_ReadsFromConsoleIn()
    {
        var validJson = """[["ACD Full Term", 0.26], ["RI 3yr", 0.40]]""";

        var originalIn = Console.In;
        Console.SetIn(new StringReader(validJson));
        try
        {
            var result = Deserialize("-");

            Assert.Equal(2, result.Count);
            Assert.Equal(2, result[0].Count);
            // Values come back as JsonElement when deserializing List<List<object?>>
            var firstLabel = result[0][0] is JsonElement je ? je.GetString() : result[0][0]?.ToString();
            Assert.Equal("ACD Full Term", firstLabel);
        }
        finally
        {
            Console.SetIn(originalIn);
        }
    }
}
