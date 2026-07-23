using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Tests backward compatibility of CLI parameter aliases.
/// Before code generation, CLI used short parameter names like --sheet and --range.
/// After generator refactoring, kebab-case names are generated from Core camelCase (--sheet-name, --range-address).
/// These tests verify that short aliases continue to work for backward compatibility.
/// Bug context: User workflows broke with "sheetName is required" error when using --sheet and --range.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public class ParameterAliasBackwardCompatTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _testFile;

    public ParameterAliasBackwardCompatTests(ITestOutputHelper output)
    {
        _output = output;
        // Use a temp file path - session open will create it if it doesn't exist
        _testFile = Path.Combine(Path.GetTempPath(), $"ParamAlias_{Guid.NewGuid():N}.xlsx");
    }

    [Fact]
    public async Task RangeSetValues_ShortAliases_WorksWithoutError()
    {
        // Create test file
        Sbroenne.ExcelMcp.ComInterop.Session.ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) => 0,
            CancellationToken.None);

        // Open session via CLI
        var openResult = await CliProcessHelper.RunAsync($"session open {_testFile}");
        Assert.Equal(0, openResult.ExitCode);
        var openJson = JsonDocument.Parse(openResult.Stdout);
        var sid = openJson.RootElement.GetProperty("sessionId").GetString();

        try
        {
            // Act - Use OLD short parameter names: --sheet and --range (pre-generator syntax)
            var result = await CliProcessHelper.RunAsync(
                $"range set-values --session {sid} --sheet Sheet1 --range A1 --values \"[[\\\"BackwardCompat\\\"]]\"");

            // Assert - Should accept short aliases without error
            _output.WriteLine($"Exit code: {result.ExitCode}");
            _output.WriteLine($"Stdout: {result.Stdout}");
            _output.WriteLine($"Stderr: {result.Stderr}");
            Assert.Equal(0, result.ExitCode);

            var json = JsonDocument.Parse(result.Stdout);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean(),
                "CLI should accept --sheet and --range aliases");
        }
        finally
        {
            await CliProcessHelper.RunAsync($"session close --session {sid} --save false");
        }
    }

    [Fact]
    public async Task RangeSetValues_LongNames_WorksWithoutError()
    {
        // Create test file
        Sbroenne.ExcelMcp.ComInterop.Session.ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) => 0,
            CancellationToken.None);

        // Open session via CLI
        var openResult = await CliProcessHelper.RunAsync($"session open {_testFile}");
        Assert.Equal(0, openResult.ExitCode);
        var openJson = JsonDocument.Parse(openResult.Stdout);
        var sid = openJson.RootElement.GetProperty("sessionId").GetString();

        try
        {
            // Act - Use NEW long parameter names (generated syntax)
            var result = await CliProcessHelper.RunAsync(
                $"range set-values --session {sid} --sheet-name Sheet1 --range-address B1 --values \"[[\\\"NewSyntax\\\"]]\"");

            // Assert
            _output.WriteLine($"Exit code: {result.ExitCode}");
            _output.WriteLine($"Stdout: {result.Stdout}");
            Assert.Equal(0, result.ExitCode);

            var json = JsonDocument.Parse(result.Stdout);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean(),
                "CLI should accept --sheet-name and --range-address");
        }
        finally
        {
            await CliProcessHelper.RunAsync($"session close --session {sid} --save false");
        }
    }

    [Fact]
    public async Task RangeGetValues_ShortAliases_WorksWithoutError()
    {
        // Create test file
        Sbroenne.ExcelMcp.ComInterop.Session.ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) => 0,
            CancellationToken.None);

        // Open session via CLI
        var openResult = await CliProcessHelper.RunAsync($"session open {_testFile}");
        Assert.Equal(0, openResult.ExitCode);
        var openJson = JsonDocument.Parse(openResult.Stdout);
        var sid = openJson.RootElement.GetProperty("sessionId").GetString();

        try
        {
            // Setup - Write a value first
            await CliProcessHelper.RunAsync(
                $"range set-values --session {sid} --sheet-name Sheet1 --range-address D1 --values \"[[\\\"ReadTest\\\"]]\"");

            // Act - Read using OLD short parameter names
            var result = await CliProcessHelper.RunAsync(
                $"range get-values --session {sid} --sheet Sheet1 --range D1");

            // Assert
            _output.WriteLine($"Exit code: {result.ExitCode}");
            _output.WriteLine($"Stdout: {result.Stdout}");
            Assert.Equal(0, result.ExitCode);

            var json = JsonDocument.Parse(result.Stdout);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean(),
                "CLI should accept --sheet and --range aliases for get-values");
        }
        finally
        {
            await CliProcessHelper.RunAsync($"session close --session {sid} --save false");
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
            try { File.Delete(_testFile); }
            catch { /* Best effort cleanup */ }
        }
        GC.SuppressFinalize(this);
    }
}
