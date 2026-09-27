using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacVbaHelperProtocolTests
{
    [Fact]
    public void Request_RoundTripsVersionIdentityPathActionAndArguments()
    {
        var request = MacVbaHelperProtocol.CreateRequest(
            "0123456789abcdef0123456789abcdef",
            "/tmp/exact workbook.xlsm",
            "vba.view",
            new { moduleName = "Module1" });

        var parsed = MacVbaHelperProtocol.ParseRequest(request);

        Assert.Equal(MacVbaHelperProtocol.Version, parsed.Version);
        Assert.Equal("0123456789abcdef0123456789abcdef", parsed.RequestId);
        Assert.Equal("/tmp/exact workbook.xlsm", parsed.WorkbookPath);
        Assert.Equal("vba.view", parsed.Action);
        Assert.Equal("Module1", parsed.Arguments.GetProperty("moduleName").GetString());
    }

    [Theory]
    [InlineData("analysis.create-scenario")]
    [InlineData("analysis.show-scenario")]
    public void Request_AllowsFixedScenarioHelperActions(string action)
    {
        var request = MacVbaHelperProtocol.CreateRequest(
            "0123456789abcdef0123456789abcdef",
            "/tmp/exact workbook.xlsx",
            action,
            new { sheetName = "Inputs", scenarioName = "Baseline" });

        Assert.Equal(action, MacVbaHelperProtocol.ParseRequest(request).Action);
    }

    [Theory]
    [InlineData("unknown.action")]
    [InlineData("vba.run")]
    [InlineData("helper.eval")]
    [InlineData("powerquery.refresh")]
    [InlineData("powerquery.evaluate")]
    public void Request_RejectsActionsOutsideFixedAllowlist(string action)
    {
        var error = Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.CreateRequest(
            "0123456789abcdef0123456789abcdef",
            "/tmp/workbook.xlsm",
            action,
            new { }));

        Assert.Contains("not allowed", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Request_PreservesExplicitNullArguments()
    {
        var request = MacVbaHelperProtocol.CreateRequest(
            "0123456789abcdef0123456789abcdef",
            "/tmp/exact workbook.xlsx",
            "analysis.create-scenario",
            new
            {
                sheetName = "Inputs",
                scenarioName = "Baseline",
                changingCells = "A1",
                values = new object?[] { 1 },
                comment = (string?)null,
                locked = true,
                hidden = false
            });

        Assert.Equal(
            JsonValueKind.Null,
            MacVbaHelperProtocol.ParseRequest(request).Arguments.GetProperty("comment").ValueKind);
    }

    [Theory]
    [InlineData("")]
    [InlineData("request-1")]
    [InlineData("0123456789ABCDEF0123456789ABCDEF")]
    public void Request_RequiresCanonicalRequestId(string requestId)
    {
        Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.CreateRequest(
            requestId,
            "/tmp/workbook.xlsm",
            "helper.capabilities",
            new { }));
    }

    [Fact]
    public void Request_RejectsOversizedPayload()
    {
        var source = new string('x', MacVbaHelperProtocol.MaxPayloadBytes);

        Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.CreateRequest(
            "0123456789abcdef0123456789abcdef",
            "/tmp/workbook.xlsm",
            "vba.import",
            new { moduleName = "Module1", source }));
    }

    [Fact]
    public void Response_RequiresMatchingCorrelationAndSuccessShape()
    {
        const string requestId = "0123456789abcdef0123456789abcdef";
        var result = MacVbaHelperProtocol.ParseResponse(
            """
            {"version":1,"requestId":"0123456789abcdef0123456789abcdef","success":true,"result":{"helperVersion":"1.0.0"},"error":null}
            """,
            requestId);

        Assert.True(result.Success);
        Assert.Equal("1.0.0", result.Result!.Value.GetProperty("helperVersion").GetString());

        Assert.Throws<InvalidOperationException>(() => MacVbaHelperProtocol.ParseResponse(
            """
            {"version":1,"requestId":"ffffffffffffffffffffffffffffffff","success":true,"result":{},"error":null}
            """,
            requestId));
        Assert.Throws<InvalidOperationException>(() => MacVbaHelperProtocol.ParseResponse(
            """
            {"version":1,"requestId":"0123456789abcdef0123456789abcdef","success":true,"result":{},"error":{"category":"ComInterop","code":"bad","message":"bad"}}
            """,
            requestId));
    }

    [Fact]
    public void Response_RejectsDuplicateErrorPropertiesAndOversizedPayload()
    {
        const string requestId = "0123456789abcdef0123456789abcdef";
        Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.ParseResponse(
            """
            {"version":1,"requestId":"0123456789abcdef0123456789abcdef","success":false,"result":null,"error":{"category":"ComInterop","code":"bad","code":"duplicate","message":"bad"}}
            """,
            requestId));

        var oversized = $$"""
            {"version":1,"requestId":"{{requestId}}","success":true,"result":{"value":"{{new string('x', MacVbaHelperProtocol.MaxPayloadBytes)}}"},"error":null}
            """;
        Assert.Throws<ArgumentException>(() =>
            MacVbaHelperProtocol.ParseResponse(oversized, requestId));
    }

    [Fact]
    public void EmbeddedHelper_IsOriginalFixedDispatcherWithNoArbitraryEvaluation()
    {
        using var stream = typeof(MacVbaHelperProtocol).Assembly.GetManifestResourceStream(
            "Sbroenne.ExcelMcp.Service.Mac.ExcelMcpHelper.bas");
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream!);
        var source = reader.ReadToEnd();

        Assert.Contains("Public Function ExcelMcpDispatch(ByVal requestJson As String) As String", source);
        Assert.Contains("candidate.FullName", source);
        Assert.Contains("VBProject.VBComponents", source);
        Assert.Contains("CodeModule.AddFromString", source);
        Assert.Contains("Queries.Add", source);
        Assert.Contains("QueryTable.WorkbookConnection", source);
        Assert.Contains("scenarios.Add", source);
        Assert.Contains("scenario.Show", source);
        Assert.Contains("scenarioCreateShow", source);
        Assert.Contains("engineCapabilities", source);
        Assert.Contains("helper_target_forbidden", source);
        Assert.Contains("If cleanupNumber <> 0 Then", source);
        Assert.Contains("If rollbackNumber <> 0 Then", source);
        Assert.Contains(@"""rollback_failed""", source);
        Assert.DoesNotContain("Application.Evaluate", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ExecuteGlobal", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("VBComponents.Import", source, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EmbeddedHelper_RespectsVbaStatementContinuationLimit()
    {
        using var stream = typeof(MacVbaHelperProtocol).Assembly.GetManifestResourceStream(
            "Sbroenne.ExcelMcp.Service.Mac.ExcelMcpHelper.bas");
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream!);
        var lines = reader.ReadToEnd().ReplaceLineEndings("\n").Split('\n');
        var continuations = 0;
        foreach (var line in lines)
        {
            if (line.TrimEnd().EndsWith(" _", StringComparison.Ordinal))
            {
                continuations++;
                Assert.True(
                    continuations <= 24,
                    $"VBA statement ending near '{line.Trim()}' exceeds 24 continuations.");
            }
            else
            {
                continuations = 0;
            }
        }
    }

    [Fact]
    public void BuildOutput_IncludesReviewableBootstrapSource()
    {
        var sourcePath = Path.Combine(
            AppContext.BaseDirectory,
            "helpers",
            "ExcelMcpHelper.bas");

        Assert.True(File.Exists(sourcePath), $"Expected helper source at '{sourcePath}'.");
        Assert.Contains(
            "Public Function ExcelMcpDispatch",
            File.ReadAllText(sourcePath),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Request_RejectsMalformedDuplicateAndUnknownProperties()
    {
        Assert.ThrowsAny<JsonException>(() => MacVbaHelperProtocol.ParseRequest("{"));
        Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.ParseRequest(
            """
            {"version":1,"version":1,"requestId":"0123456789abcdef0123456789abcdef","workbookPath":"/tmp/a.xlsm","action":"helper.capabilities","arguments":{}}
            """));
        Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.ParseRequest(
            """
            {"version":1,"requestId":"0123456789abcdef0123456789abcdef","workbookPath":"/tmp/a.xlsm","action":"helper.capabilities","arguments":{},"extra":true}
            """));
    }

    [Fact]
    public void Utf8Limit_CountsNonAsciiBytesWithoutTruncation()
    {
        var source = new string('\u20ac', MacVbaHelperProtocol.MaxPayloadBytes / 3);

        var error = Assert.Throws<ArgumentException>(() => MacVbaHelperProtocol.CreateRequest(
            "0123456789abcdef0123456789abcdef",
            "/tmp/workbook.xlsm",
            "vba.import",
            new { moduleName = "Module1", source }));

        Assert.Contains("UTF-8 bytes", error.Message, StringComparison.Ordinal);
        Assert.Contains("262144", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HelperDispatchScript_UsesOnlyFixedMacroAndSingleJsonArgument()
    {
        var script = MacAutomationHost.CreateHelperDispatchScriptForTests(
            "/Users/test/Library/Application Support/ExcelMcp/ExcelMcpHelper.xlam",
            """{"version":1,"requestId":"0123456789abcdef0123456789abcdef"}""");

        Assert.Contains(
            "/Users/test/Library/Application Support/ExcelMcp/ExcelMcpHelper.xlam",
            script);
        Assert.Contains("ExcelMcpHelper.xlam!ExcelMcpDispatch", script);
        Assert.Contains("arg1", script);
        Assert.Contains("Exactly one ExcelMcpHelper.xlam must be open.", script);
        Assert.DoesNotContain("macroName", script, StringComparison.OrdinalIgnoreCase);
    }
}
