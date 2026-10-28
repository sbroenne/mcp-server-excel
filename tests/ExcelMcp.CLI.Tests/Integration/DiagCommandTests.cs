using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for the diag CLI command.
/// Tests the full CLI pipeline end-to-end: process launch → service dispatch → JSON response.
/// These tests do NOT require Excel — they validate CLI infrastructure (parameter parsing,
/// validation, routing, JSON serialization, exit codes).
/// </summary>
[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Diag")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
[Trait("AdapterTestKind", "System")]
public sealed class DiagCommandTests
{
    private readonly ITestOutputHelper _output;

    public DiagCommandTests(ITestOutputHelper output) => _output = output;

    // ============================================
    // PING - Basic connectivity
    // ============================================

    [Fact]
    public async Task Ping_ReturnsSuccessWithoutSession()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag ping");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("diag.ping", json.RootElement.GetProperty("command").GetString());
        Assert.Equal("pong", json.RootElement.GetProperty("data").GetProperty("message").GetString());
        Assert.True(json.RootElement.GetProperty("data").TryGetProperty("timestamp", out _));
    }

    // ============================================
    // ECHO - Required string parameter
    // ============================================

    [Fact]
    public async Task Echo_WithMessage_ReturnsEchoedMessage()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"Hello World\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("diag.echo", json.RootElement.GetProperty("command").GetString());
        Assert.Equal("Hello World", json.RootElement.GetProperty("data").GetProperty("message").GetString());
    }

    [Fact]
    public async Task Echo_WithMessageAndTag_ReturnsBoth()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"test\" --tag \"v1\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("test", json.RootElement.GetProperty("data").GetProperty("message").GetString());
        Assert.Equal("v1", json.RootElement.GetProperty("data").GetProperty("tag").GetString());
    }

    [Fact]
    public async Task Echo_WithoutMessage_ReturnsValidationError()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo");
        _output.WriteLine(result.Stdout);

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("message", json.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.Contains("required", json.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task Echo_WithSpecialCharacters_PreservesMessage()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"hello & goodbye <world>\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.Equal("hello & goodbye <world>", json.RootElement.GetProperty("data").GetProperty("message").GetString());
    }

    [Fact]
    public async Task Echo_WithUnknownOption_ReturnsValidationError()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag echo --message \"sentinel\" --totally-bogus-flag xyz");

        Assert.NotEqual(0, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("totally-bogus-flag", json.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task SlicerCreate_WithoutRequiredParameters_ReportsAllMissingParameters()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "slicer create-slicer --session unused");

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());

        var error = json.RootElement.GetProperty("error").GetProperty("message").GetString();
        Assert.Contains("pivotTableName", error);
        Assert.Contains("fieldName", error);
        Assert.Contains("slicerName", error);
        Assert.Contains("destinationSheet", error);
        Assert.Contains("position", error);
    }

    // ============================================
    // VALIDATE-PARAMS - Multiple parameter types
    // ============================================

    [Fact]
    public async Task ValidateParams_AllParams_ReturnsAll()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --name \"test\" --count 42 --label \"optional\" --verbose true");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("diag.validate-params", json.RootElement.GetProperty("command").GetString());

        var parameters = json.RootElement.GetProperty("data").GetProperty("parameters");
        Assert.Equal("test", parameters.GetProperty("name").GetString());
        Assert.Equal(42, parameters.GetProperty("count").GetInt32());
        Assert.Equal("optional", parameters.GetProperty("label").GetString());
        Assert.True(parameters.GetProperty("verbose").GetBoolean());
    }

    [Fact]
    public async Task ValidateParams_RequiredOnly_DefaultsApplied()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --name \"test\" --count 5");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());

        var parameters = json.RootElement.GetProperty("data").GetProperty("parameters");
        Assert.Equal("test", parameters.GetProperty("name").GetString());
        Assert.Equal(5, parameters.GetProperty("count").GetInt32());
        Assert.Null(parameters.GetProperty("label").GetString());
        Assert.False(parameters.GetProperty("verbose").GetBoolean());
    }

    [Fact]
    public async Task ValidateParams_MissingName_ReturnsValidationError()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --count 5");
        _output.WriteLine(result.Stdout);

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("name", json.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.Contains("required", json.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task ValidateParams_MissingCount_ReturnsValidationError()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --name \"test\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stdout: {result.Stdout}");

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("count", json.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.Contains("required", json.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    // ============================================
    // INVALID ACTION - Unknown action handling
    // ============================================

    [Fact]
    public async Task InvalidAction_ReturnsError()
    {
        var result = await CliProcessHelper.RunAsync("diag nonexistent");
        _output.WriteLine(result.Stdout);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("Invalid action", result.Stdout);
    }

    // ============================================
    // JSON FORMAT - Output structure validation
    // ============================================

    [Fact]
    public async Task ErrorResponse_HasCorrectStructure()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo");

        Assert.DoesNotContain("Unhandled error", result.Stdout);
        Assert.True(json.RootElement.TryGetProperty("success", out var success));
        Assert.False(success.GetBoolean());
        Assert.True(json.RootElement.GetProperty("error").TryGetProperty("message", out var error));
        Assert.False(string.IsNullOrWhiteSpace(error.GetString()));
        Assert.False(json.RootElement.TryGetProperty("errorMessage", out _));
    }

    [Fact]
    public async Task SuccessResponse_HasCorrectStructure()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag ping");

        // Success responses must have 'success' property
        Assert.DoesNotContain("Unhandled error", result.Stdout);
        Assert.True(json.RootElement.TryGetProperty("success", out var success));
        Assert.True(success.GetBoolean());

        Assert.False(json.RootElement.TryGetProperty("error", out _),
            "Successful responses must not include an error.");
    }
}
