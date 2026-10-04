using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class GeneratedActionContractProtocolTests : McpIntegrationTestBase
{
    public GeneratedActionContractProtocolTests(ITestOutputHelper output)
        : base(output, "GeneratedActionContractProtocolClient")
    {
    }

    [Theory]
    [InlineData("calculation_mode", "calculate", "mode", "mode", "manual")]
    [InlineData("powerquery", "delete", "m_code", "mCode", "let Source = 1 in Source")]
    [InlineData("powerquery", "delete", "m_code_file", "mCodeFile", "missing-query.pq")]
    public async Task ToolCall_RejectsParameterFromAnotherActionWithStructuredJson(
        string toolName,
        string action,
        string invalidParameter,
        string expectedErrorParameter,
        string invalidValue)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "missing-session",
            [invalidParameter] = invalidValue
        };
        if (action == "calculate")
        {
            arguments["scope"] = "application";
        }
        else
        {
            arguments["query_name"] = "Probe";
        }

        var result = await CallToolAsync(toolName, arguments);

        using var document = ParseJsonResult(result, $"{toolName}.{action}");
        AssertFailureEnvelope(
            document.RootElement,
            $"{toolName}.{action}",
            nameof(ArgumentException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(expectedErrorParameter, document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ToolCall_RejectsExplicitNullParameterFromAnotherAction()
    {
        var result = await CallToolAsync(
            "calculation_mode",
            new Dictionary<string, object?>
            {
                ["action"] = "calculate",
                ["session_id"] = "missing-session",
                ["scope"] = "application",
                ["mode"] = null
            });

        using var document = ParseJsonResult(result, "calculation.calculate");
        AssertFailureEnvelope(
            document.RootElement,
            "calculation.calculate",
            nameof(ArgumentException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(
            "mode",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ToolCall_AllowEmptyRequiredStringRejectsExplicitNullBeforeDispatch()
    {
        var result = await CallToolAsync(
            "conditionalformat",
            new Dictionary<string, object?>
            {
                ["action"] = "clear-rules",
                ["session_id"] = "missing-session",
                ["sheet_name"] = null,
                ["range_address"] = "A1"
            });

        using var document = ParseJsonResult(result, "conditionalformat.clear-rules");
        AssertFailureEnvelope(
            document.RootElement,
            "conditionalformat.clear-rules",
            nameof(ArgumentException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(
            "sheetName",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(42)]
    [InlineData(true)]
    public async Task ToolCall_AllowEmptyRequiredStringRejectsNonStringBeforeDispatch(
        object invalidValue)
    {
        var result = await CallToolAsync(
            "conditionalformat",
            new Dictionary<string, object?>
            {
                ["action"] = "clear-rules",
                ["session_id"] = "missing-session",
                ["sheet_name"] = invalidValue,
                ["range_address"] = "A1"
            });

        Assert.Contains("error", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("conditionalformat", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ToolCall_AllowEmptyRequiredStringAcceptsEmptyStringBeforeDispatch()
    {
        var result = await CallToolAsync(
            "conditionalformat",
            new Dictionary<string, object?>
            {
                ["action"] = "clear-rules",
                ["session_id"] = "missing-session",
                ["sheet_name"] = string.Empty,
                ["range_address"] = "A1"
            });

        using var document = ParseJsonResult(result, "conditionalformat.clear-rules");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            "sheetName",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public async Task ToolCall_RejectsBlankWorksheetNameBeforeSessionDispatch(string sheetName)
    {
        var result = await CallToolAsync(
            "worksheet",
            new Dictionary<string, object?>
            {
                ["action"] = "create",
                ["session_id"] = "missing-session",
                ["sheet_name"] = sheetName
            });

        using var document = ParseJsonResult(result, "worksheet.create");
        AssertFailureEnvelope(
            document.RootElement,
            "worksheet.create",
            nameof(ArgumentException),
            expectedErrorCategory: "InvalidInput");
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("sheetName", error, StringComparison.Ordinal);
        Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);
    }


    [Theory]
    [InlineData("powerquery", "load-to", "load_destination", "loadDestination", "not-a-destination")]
    [InlineData("powerquery", "load-to", "load_destination", "loadDestination", "work_sheet")]
    [InlineData("powerquery", "load-to", "load_destination", "loadDestination", "work-sheet")]
    [InlineData("powerquery", "load-to", "load_destination", "loadDestination", "0")]
    [InlineData("chart", "create-from-range", "chart_type", "chartType", "not-a-chart")]
    public async Task ToolCall_RejectsUnknownEnumWithStructuredJson(
        string toolName,
        string action,
        string enumParameter,
        string expectedErrorParameter,
        string invalidValue)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "missing-session",
            [enumParameter] = invalidValue
        };
        if (toolName == "powerquery")
        {
            arguments["query_name"] = "Probe";
        }
        else
        {
            arguments["sheet_name"] = "Model";
            arguments["source_range_address"] = "A1:B2";
        }

        var result = await CallToolAsync(toolName, arguments);

        using var document = ParseJsonResult(result, $"{toolName}.{action}");
        AssertFailureEnvelope(
            document.RootElement,
            $"{toolName}.{action}",
            nameof(ArgumentException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(expectedErrorParameter, document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains(invalidValue, document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("worksheet")]
    [InlineData("WORKSHEET")]
    public async Task ToolCall_AcceptsExactAliasIgnoringCase(string loadDestination)
    {
        var result = await CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "load-to",
                ["session_id"] = "missing-session",
                ["query_name"] = "Probe",
                ["load_destination"] = loadDestination
            });

        using var document = ParseJsonResult(result, "powerquery.load-to");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            "Invalid value",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PowerQueryLoadTo_RejectsTimeoutSecondsAsActionInapplicable()
    {
        var result = await CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "load-to",
                ["session_id"] = "missing-session",
                ["query_name"] = "Probe",
                ["load_destination"] = "worksheet",
                ["timeout_seconds"] = 60
            });

        using var document = ParseJsonResult(result, "powerquery.load-to");
        AssertFailureEnvelope(
            document.RootElement,
            "powerquery.load-to",
            nameof(ArgumentException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(
            "timeout",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FileCreate_ForwardsMacroEnabledThroughSessionProtocol()
    {
        var path = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        await File.WriteAllTextAsync(path, string.Empty);
        try
        {
            var result = await CallToolAsync(
                "file",
                new Dictionary<string, object?>
                {
                    ["action"] = "create",
                    ["path"] = path
                });

            using var document = ParseJsonResult(result, "file.create");
            Assert.False(document.RootElement.GetProperty("success").GetBoolean());
            var error = document.RootElement.GetProperty("errorMessage").GetString();
            Assert.Contains("already exists", error, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("macroEnabled", error, StringComparison.Ordinal);
            Assert.DoesNotContain("Unknown parameter", error, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task ToolCall_RejectsConflictingInlineAndFileInputs()
    {
        var path = Path.GetTempFileName();
        try
        {
            var result = await CallToolAsync(
                "powerquery",
                new Dictionary<string, object?>
                {
                    ["action"] = "evaluate",
                    ["session_id"] = "missing-session",
                    ["m_code"] = "let Source = 1 in Source",
                    ["m_code_file"] = path
                });

            using var document = ParseJsonResult(result, "powerquery.evaluate");
            AssertFailureEnvelope(
                document.RootElement,
                "powerquery.evaluate",
                nameof(ArgumentException),
                expectedErrorCategory: "InvalidInput");
            var error = document.RootElement.GetProperty("errorMessage").GetString();
            Assert.Contains("mCode", error, StringComparison.Ordinal);
            Assert.Contains("mCodeFile", error, StringComparison.Ordinal);
            Assert.DoesNotContain("session", error, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("connection", "refresh", "connection_name")]
    [InlineData("datamodel", "refresh", null)]
    [InlineData("pivottable", "refresh", "pivot_table_name")]
    [InlineData("powerquery", "refresh", "query_name")]
    [InlineData("vba", "run", "procedure_name")]
    public async Task ToolCall_AcceptsIntegerTimeoutSecondsAcrossGeneratedCategories(
        string toolName,
        string action,
        string? requiredName)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "missing-session",
            ["timeout_seconds"] = 60
        };
        if (requiredName != null)
        {
            arguments[requiredName] = "Probe";
        }

        var result = await CallToolAsync(toolName, arguments);

        using var document = ParseJsonResult(result, $"{toolName}.{action}");
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            "timeout",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("powerquery", "refresh", "query_name", -1)]
    [InlineData("connection", "refresh", "connection_name", 0)]
    [InlineData("datamodel", "refresh", null, 2147484)]
    [InlineData("pivottable", "refresh", "pivot_table_name", 0)]
    [InlineData("vba", "run", "procedure_name", 0)]
    public async Task ToolCall_RejectsOutOfRangeTimeoutBeforeSessionLookup(
        string toolName,
        string action,
        string? requiredName,
        int timeoutSeconds)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["session_id"] = "missing-session",
            ["timeout_seconds"] = timeoutSeconds
        };
        if (requiredName != null)
        {
            arguments[requiredName] = "Probe";
        }

        var result = await CallToolAsync(toolName, arguments);

        using var document = ParseJsonResult(result, $"{toolName}.{action}");
        AssertFailureEnvelope(
            document.RootElement,
            $"{toolName}.{action}",
            nameof(ArgumentOutOfRangeException),
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(
            "timeout",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            "session",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }

}
