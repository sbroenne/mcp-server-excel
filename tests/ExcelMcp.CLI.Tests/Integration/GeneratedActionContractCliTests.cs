using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class GeneratedActionContractCliTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDirectory;

    public GeneratedActionContractCliTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDirectory = Path.Join(Path.GetTempPath(), $"GeneratedActionContractCliTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDirectory);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDirectory, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    [Theory]
    [InlineData(
        "calculationmode calculate --session missing-session --scope workbook --mode manual",
        "mode",
        "calculate")]
    [InlineData(
        "powerquery delete --session missing-session --query-name Probe --m-code \"let Source = 1 in Source\"",
        "mCode",
        "delete")]
    [InlineData(
        "powerquery load-to --session missing-session --query-name Probe --load-destination not-a-destination",
        "loadDestination",
        "not-a-destination")]
    [InlineData(
        "powerquery load-to --session missing-session --query-name Probe --load-destination work_sheet",
        "loadDestination",
        "work_sheet")]
    [InlineData(
        "powerquery load-to --session missing-session --query-name Probe --load-destination work-sheet",
        "loadDestination",
        "work-sheet")]
    [InlineData(
        "powerquery load-to --session missing-session --query-name Probe --load-destination worksheet --timeout 30",
        "timeout",
        "load-to")]
    [InlineData(
        "powerquery refresh --session missing-session --query-name Probe --timeout -1",
        "timeout",
        "2147483")]
    [InlineData(
        "connection refresh --session missing-session --connection-name Probe --timeout 0",
        "timeout",
        "2147483")]
    [InlineData(
        "chart create-from-range --session missing-session --sheet Model --source-range-address A1:B2 --chart-type 999",
        "chartType",
        "999")]
    [InlineData(
        "chart create-from-range --session missing-session --sheet Model --source-range-address A1:B2 --chart-type 51",
        "chartType",
        "51")]
    [InlineData(
        "chart create-from-range --session missing-session --sheet Model --source-range-address A1:B2",
        "chartType",
        "required")]
    public async Task DirectCommand_RejectsInvalidGeneratedContractBeforeDaemonDispatch(
        string arguments,
        string expectedParameter,
        string expectedDetail)
    {
        var result = await CliProcessHelper.RunAsync(arguments);
        _output.WriteLine($"stdout: {result.Stdout}");
        _output.WriteLine($"stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
        var combinedOutput = result.Stdout + result.Stderr;
        Assert.Contains(expectedParameter, combinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(expectedDetail, combinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", combinedOutput, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(
        """{"command":"calculation.calculate","sessionId":"missing-session","args":{"scope":"workbook","mode":"manual"}}""",
        "mode",
        "calculate")]
    [InlineData(
        """{"command":"powerquery.load-to","sessionId":"missing-session","args":{"queryName":"Probe","loadDestination":"not-a-destination"}}""",
        "loadDestination",
        "not-a-destination")]
    [InlineData(
        """{"command":"powerquery.load-to","sessionId":"missing-session","args":{"queryName":"Probe","loadDestination":"0"}}""",
        "loadDestination",
        "0")]
    [InlineData(
        """{"command":"powerquery.load-to","sessionId":"missing-session","args":{"queryName":"Probe","loadDestination":"work_sheet"}}""",
        "loadDestination",
        "work_sheet")]
    [InlineData(
        """{"command":"powerquery.load-to","sessionId":"missing-session","args":{"queryName":"Probe","loadDestination":"work-sheet"}}""",
        "loadDestination",
        "work-sheet")]
    [InlineData(
        """{"command":"calculation.calculate","sessionId":"missing-session","args":{"scope":5}}""",
        "scope",
        "String")]
    [InlineData(
        """{"command":"chart.create-from-range","sessionId":"missing-session","args":{"sheetName":"Model","sourceRangeAddress":"A1:B2"}}""",
        "chartType",
        "required")]
    [InlineData(
        """{"command":"powerquery.refresh","sessionId":"missing-session","args":{"queryName":"Probe","timeout":-1}}""",
        "timeout",
        "2147483")]
    [InlineData(
        """{"command":"connection.refresh","sessionId":"missing-session","args":{"connectionName":"Probe","timeout":0}}""",
        "timeout",
        "2147483")]
    [InlineData(
        """{"command":"powerquery.refresh","sessionId":"missing-session","args":{"queryName":"Probe","timeout":"600"}}""",
        "timeout",
        "Int32")]
    public async Task RawBatch_RejectsInvalidGeneratedContractBeforeDaemonDispatch(
        string entryJson,
        string expectedParameter,
        string expectedDetail)
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(inputPath, $"[{entryJson}]");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);
        _output.WriteLine($"stdout: {result.Stdout}");
        _output.WriteLine($"stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout.Trim());
        var error = output.RootElement.GetProperty("error").GetString();
        Assert.Contains(expectedParameter, error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(expectedDetail, error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("42")]
    [InlineData("true")]
    public async Task RawBatch_AllowEmptyRequiredStringRejectsNonStringBeforeDispatch(
        string jsonValue)
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            $"[{{\"command\":\"conditionalformat.clear-rules\",\"sessionId\":\"missing-session\"," +
            $"\"args\":{{\"sheetName\":{jsonValue},\"rangeAddress\":\"A1\"}}}}]");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("sheetName", result.Stdout + result.Stderr, StringComparison.Ordinal);
        Assert.Contains("JSON string", result.Stdout + result.Stderr, StringComparison.Ordinal);
        Assert.DoesNotContain("session", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RawBatch_AllowEmptyRequiredStringAcceptsEmptyStringBeforeDispatch()
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            """
            [{"command":"conditionalformat.clear-rules","sessionId":"missing-session","args":{"sheetName":"","rangeAddress":"A1"}}]
            """);

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("session", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("sheetName", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RawBatch_ReportsMissingFileAliasesAsIndexedValidationErrors()
    {
        var firstMissingPath = Path.Join(_tempDirectory, "missing-first.m");
        var secondMissingPath = Path.Join(_tempDirectory, "missing-second.m");
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            JsonSerializer.Serialize(new object[]
            {
                new
                {
                    command = "powerquery.evaluate",
                    args = new { mCodeFile = firstMissingPath }
                },
                new
                {
                    command = "powerquery.evaluate",
                    args = new { mCodeFile = secondMissingPath }
                }
            }));

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        var outputLines = result.Stdout.Split(
            Environment.NewLine,
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        Assert.Equal(2, outputLines.Length);
        for (var index = 0; index < outputLines.Length; index++)
        {
            using var output = JsonDocument.Parse(outputLines[index]);
            Assert.Equal(index, output.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("powerquery.evaluate", output.RootElement.GetProperty("command").GetString());
            Assert.False(output.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("not found", output.RootElement.GetProperty("error").GetString(), StringComparison.OrdinalIgnoreCase);
        }
    }

    [Theory]
    [InlineData(
        "powerquery load-to --session missing-session --query-name Probe --load-destination worksheet")]
    [InlineData(
        "powerquery load-to --session missing-session --query-name Probe --load-destination WORKSHEET")]
    public async Task DirectCommand_AcceptsExactAliasIgnoringCase(string arguments)
    {
        var result = await CliProcessHelper.RunAsync(arguments);

        Assert.Equal(1, result.ExitCode);
        var combinedOutput = result.Stdout + result.Stderr;
        Assert.Contains("session", combinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Invalid value", combinedOutput, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("worksheet")]
    [InlineData("WORKSHEET")]
    public async Task RawBatch_AcceptsExactAliasIgnoringCase(string loadDestination)
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(inputPath, JsonSerializer.Serialize(new[]
        {
            new
            {
                command = "powerquery.load-to",
                sessionId = "missing-session",
                args = new { queryName = "Probe", loadDestination }
            }
        }));

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout.Trim());
        var error = output.RootElement.GetProperty("error").GetString();
        Assert.Contains("session", error, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Invalid value", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("create", 9)]
    [InlineData("create", 3601)]
    [InlineData("open", 9)]
    [InlineData("open", 3601)]
    public async Task ManualSessionCommand_RejectsTimeoutOutsideDocumentedRange(
        string action,
        int timeoutSeconds)
    {
        var workbookPath = Path.Join(_tempDirectory, "timeout-contract.xlsx");
        var result = await CliProcessHelper.RunAsync(
            ["session", action, workbookPath, "--timeout", timeoutSeconds.ToString(CultureInfo.InvariantCulture)]);

        Assert.Equal(1, result.ExitCode);
        var combinedOutput = result.Stdout + result.Stderr;
        Assert.Contains("timeout", combinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("10", combinedOutput, StringComparison.Ordinal);
        Assert.Contains("3600", combinedOutput, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("""{"command":"session.open","args":{"filePath":"missing.xlsx","timeoutSeconds":9}}""")]
    [InlineData("""{"command":"session.create","args":{"filePath":"missing.xlsx","timeoutSeconds":3601}}""")]
    [InlineData("""{"command":"session.open","args":{"filePath":"missing.xlsx","timeoutSeconds":"120"}}""")]
    public async Task RawBatchSessionCommand_RejectsNonCanonicalTimeout(string request)
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(inputPath, $"[{request}]");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("timeout", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RawBatchSessionCommand_RejectsUnknownArgument()
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            """[{"command":"session.open","args":{"filePath":"missing.xlsx","unexpected":true}}]""");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("unexpected", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RawBatch_RejectsUnknownTopLevelProperty()
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            """[{"command":"diag.ping","args":{},"unexpected":true}]""");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("unexpected", result.Stderr + result.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(
        """{"command":"powerquery.refresh","sessionId":"missing-session","args":{"queryName":"Probe","Timeout":60}}""",
        "Timeout")]
    [InlineData(
        """{"command":"session.open","args":{"filePath":"missing.xlsx","TimeoutSeconds":120}}""",
        "TimeoutSeconds")]
    [InlineData(
        """{"command":"powerquery.evaluate","sessionId":"missing-session","args":{"MCodeFile":"query.m"}}""",
        "MCodeFile")]
    [InlineData(
        """{"command":"vba.import","sessionId":"missing-session","args":{"moduleName":"Module1","VbaCodeFile":"module.bas"}}""",
        "VbaCodeFile")]
    [InlineData(
        """{"command":"datamodel.create-measure","sessionId":"missing-session","args":{"tableName":"Sales","measureName":"Total","DaxFormulaFile":"measure.dax"}}""",
        "DaxFormulaFile")]
    [InlineData(
        """{"command":"datamodel.evaluate","sessionId":"missing-session","args":{"DaxQueryFile":"query.dax"}}""",
        "DaxQueryFile")]
    [InlineData(
        """{"command":"datamodel.execute-dmv","sessionId":"missing-session","args":{"DmvQueryFile":"query.dmv"}}""",
        "DmvQueryFile")]
    [InlineData(
        """{"command":"xmlmap.add","sessionId":"missing-session","args":{"SchemaFile":"schema.xsd"}}""",
        "SchemaFile")]
    [InlineData(
        """{"command":"xmlmap.import-xml","sessionId":"missing-session","args":{"XmlDataFile":"data.xml"}}""",
        "XmlDataFile")]
    public async Task RawBatch_RejectsNonCanonicalArgumentCasing(
        string entryJson,
        string suppliedProperty)
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(inputPath, $"[{entryJson}]");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains(suppliedProperty, result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RawBatch_RejectsNonCanonicalTopLevelCasing()
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            """[{"Command":"diag.ping","args":{}}]""");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("Command", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RawBatch_RejectsUnknownActionArgument()
    {
        var inputPath = Path.Join(_tempDirectory, $"{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            inputPath,
            """[{"command":"powerquery.refresh","sessionId":"missing-session","args":{"queryName":"Probe","timeout":60,"unexpected":true}}]""");

        var result = await CliProcessHelper.RunAsync(["batch", "--input", inputPath]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("unexpected", result.Stdout, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", result.Stdout, StringComparison.OrdinalIgnoreCase);
    }
}
