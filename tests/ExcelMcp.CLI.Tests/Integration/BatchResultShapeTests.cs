using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BatchResultShapeTests
{
    [Fact]
    public async Task Batch_SessionList_KeepsActualSessionResult()
    {
        using var service = new ExcelMcpService();
        var result = await InProcessCliHelper.RunAsync(["batch"],
            request => service.ProcessAsync(request).GetAwaiter().GetResult(),
            """{"command":"session.list"}""");

        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Empty(output.RootElement.GetProperty("data").GetProperty("sessions").EnumerateArray());
    }

    [Theory]
    [InlineData("session.list", "{}", """{"sessions":[]}""")]
    [InlineData("session.discover", "{}", """{"workbooks":[]}""")]
    [InlineData("session.notices", "{}", """{"notices":{"protectedView":false}}""")]
    [InlineData("session.test", """{"filePath":"C:\\synthetic\\book.xlsx"}""", """{"canOpen":true,"extension":".xlsx"}""")]
    public async Task Batch_ManualRead_KeepsArgumentsAndCompleteData(string command, string args, string response)
    {
        var calls = 0;
        var result = await InProcessCliHelper.RunAsync(["batch", "--session", "target"],
            request =>
            {
                calls++;
                using var sent = JsonDocument.Parse(request.Args!);
                using var expected = JsonDocument.Parse(args);
                Assert.True(JsonElement.DeepEquals(expected.RootElement, sent.RootElement));
                return new ServiceResponse { Success = true, Result = response };
            },
            JsonSerializer.Serialize(new { command, args = JsonSerializer.Deserialize<JsonElement>(args) }));
        Assert.Equal(1, calls);
        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        using var expectedData = JsonDocument.Parse(response);
        Assert.True(JsonElement.DeepEquals(expectedData.RootElement, output.RootElement.GetProperty("data")));
    }

    [Theory]
    [InlineData("powerquery.list")]
    [InlineData("session.list")]
    [InlineData("session.discover")]
    [InlineData("session.notices")]
    [InlineData("session.test")]
    public async Task Batch_RemovedFieldsUseNormalInvalidArgumentHandling(string command)
    {
        using var service = new ExcelMcpService();
        var result = await InProcessCliHelper.RunAsync(["batch"],
            request => service.ProcessAsync(request).GetAwaiter().GetResult(),
            $$$"""{"command":"{{{command}}}","args":{"fields":["name"]}}""");
        Assert.Equal(1, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.False(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", output.RootElement.GetProperty("error").GetProperty("category").GetString());
        Assert.Contains("fields", output.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Theory]
    [InlineData("service.status", """{"fields":["running"]}""")]
    [InlineData("service.shutdown", """{"fields":["running"]}""")]
    [InlineData("service.ping", """{"fields":["running"]}""")]
    [InlineData("service.status", """{"fields":null}""")]
    [InlineData("service.shutdown", """{"fields":null}""")]
    [InlineData("service.ping", """{"fields":null}""")]
    public async Task Batch_ServiceArguments_AreRejectedWithoutConnecting(string command, string args)
    {
        var result = await InProcessCliHelper.RunAsync(["batch"],
            input: $$$"""{"command":"{{{command}}}","args":{{{args}}}}""");

        Assert.Equal(1, result.ExitCode);
        Assert.Empty(result.Stderr);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal(command, output.RootElement.GetProperty("command").GetString());
        Assert.Equal(0, output.RootElement.GetProperty("index").GetInt32());
        Assert.False(output.RootElement.GetProperty("success").GetBoolean());
        var error = output.RootElement.GetProperty("error");
        Assert.Equal("InvalidInput", error.GetProperty("category").GetString());
        Assert.Contains("fields", error.GetProperty("message").GetString());
    }

    [Theory]
    [InlineData("service.status")]
    [InlineData("service.shutdown")]
    [InlineData("service.ping")]
    public async Task Batch_ServiceWithoutArguments_DispatchesNormally(string command)
    {
        var calls = 0;
        var result = await InProcessCliHelper.RunAsync(["batch"],
            request =>
            {
                calls++;
                Assert.Equal(command, request.Command);
                Assert.Equal("{}", request.Args);
                return new ServiceResponse { Success = true, Result = """{"running":true}""" };
            },
            $$$"""{"command":"{{{command}}}","args":{}}""");

        Assert.Equal(1, calls);
        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.True(output.RootElement.GetProperty("data").GetProperty("running").GetBoolean());
    }

    [Fact]
    public async Task Batch_InvalidTypePreservesOriginalCauseWithoutDispatch()
    {
        var result = await InProcessCliHelper.RunAsync(["batch"],
            input: """{"command":"calculationmode.calculate","args":{"scope":5}}""");
        Assert.Equal(1, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        var error = output.RootElement.GetProperty("error");
        Assert.Equal("ArgumentException", error.GetProperty("exceptionType").GetString());
        Assert.Contains("String", error.GetProperty("innerError").GetString());
    }

    [Fact]
    public async Task Batch_ItemTimingExcludesEarlierCommands()
    {
        var result = await InProcessCliHelper.RunAsync(["batch"],
            request =>
            {
                if (request.Command == "diag.ping")
                    Thread.Sleep(300);
                return new ServiceResponse { Success = true, Result = """{"success":true,"message":"done"}""" };
            },
            """
            {"command":"diag.ping"}
            {"command":"diag.echo","args":{"message":"fast"}}
            """);
        Assert.Equal(0, result.ExitCode);
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        using var first = JsonDocument.Parse(lines[0]);
        using var second = JsonDocument.Parse(lines[1]);
        Assert.True(first.RootElement.GetProperty("durationMs").GetInt64() >= 300);
        Assert.InRange(second.RootElement.GetProperty("durationMs").GetInt64(), 0, 250);
        Assert.Equal("diag.echo", second.RootElement.GetProperty("command").GetString());
        Assert.Equal("done", second.RootElement.GetProperty("data").GetProperty("message").GetString());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Batch_RepeatedCommands_HaveDistinctIndexes(bool arrayInput, bool fileInput)
    {
        const string first = """{"command":"range.set-values","args":{"sheetName":"Sales","rangeAddress":"A2:B3","values":[["private-value",2],[3,4]]}}""";
        const string second = """{"command":"range.set-values","args":{"values":[[5,6]],"rangeAddress":"B5:C5","sheetName":"Forecast"}}""";
        var input = arrayInput ? $"[{first},{second}]" : $"{first}\n{second}";
        var filePath = Path.Combine(Path.GetTempPath(), $"BatchSummary_{Guid.NewGuid():N}.json");
        var requests = new List<ServiceRequest>();
        try
        {
            if (fileInput)
            {
                await File.WriteAllTextAsync(filePath, input);
            }

            var result = await InProcessCliHelper.RunAsync(
                fileInput ? ["batch", "--input", filePath, "--session", "session"] : ["batch", "--session", "session"],
                request =>
                {
                    requests.Add(request);
                    return new ServiceResponse { Success = true, Result = """{"success":true}""" };
                },
                fileInput ? "" : input);

            Assert.Equal(0, result.ExitCode);
            AssertProgress(result.Stderr, 2, (1, "range.set-values", true), (2, "range.set-values", true));
            var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            Assert.Equal(2, lines.Length);
            for (var index = 0; index < lines.Length; index++)
            {
                using var json = JsonDocument.Parse(lines[index]);
                var item = json.RootElement;
                Assert.Equal(index, item.GetProperty("index").GetInt32());
                Assert.Equal("range.set-values", item.GetProperty("command").GetString());
                Assert.False(item.TryGetProperty("commandSummary", out _));
                Assert.True(item.GetProperty("durationMs").GetInt64() >= 0);
                Assert.True(item.GetProperty("success").GetBoolean());
                Assert.False(item.TryGetProperty("result", out _));
            }
            Assert.Equal(2, requests.Count);
            Assert.All(requests, request =>
            {
                Assert.Equal("session", request.SessionId);
                Assert.Equal("range.set-values", request.Command);
                Assert.Equal("cli-batch", request.Source);
                using var args = JsonDocument.Parse(request.Args!);
                Assert.False(args.RootElement.TryGetProperty("commandSummary", out _));
            });
            Assert.Contains("private-value", requests[0].Args);
        }
        finally
        {
            if (fileInput)
            {
                File.Delete(filePath);
            }
        }
    }

    [Theory]
    [InlineData("validation", false)]
    [InlineData("validation", true)]
    [InlineData("service", false)]
    [InlineData("service", true)]
    [InlineData("negative-result", false)]
    [InlineData("negative-result", true)]
    [InlineData("communication", false)]
    [InlineData("communication", true)]
    public async Task Batch_Failures_PreserveDetailsAndStopBehavior(string failure, bool stopOnError)
    {
        var first = failure == "validation"
            ? """{"command":"range.get-values","args":{"sheetName":"Sales"}}"""
            : """{"command":"range.get-values","args":{"sheetName":"Sales","rangeAddress":"A1"}}""";
        var requests = new List<ServiceRequest>();
        var result = await InProcessCliHelper.RunAsync(
            stopOnError ? ["batch", "--session", "session", "--stop-on-error"] : ["batch", "--session", "session"],
            request =>
            {
                requests.Add(request);
                if (request.Command == "diag.ping")
                {
                    return new ServiceResponse { Success = true, Result = """{"success":true}""" };
                }
                return failure switch
                {
                    "communication" => throw new IOException("Synthetic pipe failure"),
                    "negative-result" => new ServiceResponse { Success = true, Result = """{"success":false,"errorMessage":"Synthetic operation failure"}""" },
                    _ => new ServiceResponse
                    {
                        Success = false,
                        ErrorMessage = "Synthetic service failure",
                        ErrorCategory = "Synthetic",
                        ExceptionType = "SyntheticException",
                        HResult = "0x800A03EC",
                        InnerError = "Synthetic cause",
                        PartialState = """{"changed":true}"""
                    }
                };
            },
            first + "\n" + """{"command":"diag.ping"}""");

        Assert.Equal(1, result.ExitCode);
        var expectedProgress = new List<(int, string, bool)>();
        if (failure != "validation")
            expectedProgress.Add((1, "range.get-values", false));
        if (!stopOnError)
            expectedProgress.Add((2, "diag.ping", true));
        AssertProgress(result.Stderr, 2, expectedProgress.ToArray());
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(stopOnError ? 1 : 2, lines.Length);
        using var json = JsonDocument.Parse(lines[0]);
        var item = json.RootElement;
        Assert.Equal(0, item.GetProperty("index").GetInt32());
        Assert.Equal("range.get-values", item.GetProperty("command").GetString());
        Assert.False(item.TryGetProperty("commandSummary", out _));
        Assert.False(item.GetProperty("success").GetBoolean());
        var error = item.GetProperty("error");
        Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("message").GetString()));
        if (failure == "service")
        {
            Assert.Equal("Synthetic service failure", error.GetProperty("message").GetString());
            Assert.Equal("Synthetic", error.GetProperty("category").GetString());
            Assert.Equal("SyntheticException", error.GetProperty("exceptionType").GetString());
            Assert.Equal("0x800A03EC", error.GetProperty("hresult").GetString());
            Assert.Equal("Synthetic cause", error.GetProperty("innerError").GetString());
            Assert.True(item.GetProperty("partialState").GetProperty("changed").GetBoolean());
        }
        if (failure == "communication")
        {
            Assert.Equal("Communication error: Synthetic pipe failure", error.GetProperty("message").GetString());
        }
        if (failure == "negative-result")
        {
            Assert.Equal("Synthetic operation failure", error.GetProperty("message").GetString());
            Assert.False(item.TryGetProperty("result", out _));
        }
        Assert.Equal((failure == "validation" ? 0 : 1) + (stopOnError ? 0 : 1), requests.Count);
        if (!stopOnError)
        {
            using var next = JsonDocument.Parse(lines[1]);
            Assert.Equal(1, next.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("diag.ping", next.RootElement.GetProperty("command").GetString());
            Assert.True(next.RootElement.GetProperty("success").GetBoolean());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Batch_AllInvalid_ReportsIndexesWithoutConnecting(bool stopOnError)
    {
        var result = await InProcessCliHelper.RunAsync(
            stopOnError ? ["batch", "--stop-on-error"] : ["batch"],
            input: """
                {"command":"range.get-values","args":{"sheetName":"Sales"}}
                {"command":"range.get-values","args":{"sheetName":"Forecast"}}
                """);

        Assert.Equal(1, result.ExitCode);
        Assert.Empty(result.Stderr);
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(stopOnError ? 1 : 2, lines.Length);
        using var first = JsonDocument.Parse(lines[0]);
        Assert.Equal(0, first.RootElement.GetProperty("index").GetInt32());
        Assert.Equal("InvalidInput", first.RootElement.GetProperty("error").GetProperty("category").GetString());
        if (!stopOnError)
        {
            using var second = JsonDocument.Parse(lines[1]);
            Assert.Equal(1, second.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("InvalidInput", second.RootElement.GetProperty("error").GetProperty("category").GetString());
        }
    }

    [Theory]
    [InlineData(@"\uD800", true, false)]
    [InlineData(@"\uDC00", true, false)]
    [InlineData(@"\uD800x", true, false)]
    [InlineData(@"\uD800", false, false)]
    [InlineData(@"\uD800", false, true)]
    [InlineData(@"\uD800", true, true)]
    public async Task Batch_InvalidUnicodeInput_PreservesValidationErrorAndContinuation(
        string escapedName, bool includeFollowingCommand, bool stopOnError)
    {
        var original = await InProcessCliHelper.RunAsync(["batch"], input: """{"command":"nonexistent.action"}""");
        using var originalJson = JsonDocument.Parse(original.Stdout);
        var first = $$$"""{"command":"nonexistent.action","args":{"sheetName":"{{{escapedName}}}"}}""";
        var requests = new List<ServiceRequest>();
        var result = await InProcessCliHelper.RunAsync(
            stopOnError ? ["batch", "--stop-on-error"] : ["batch"],
            request =>
            {
                requests.Add(request);
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            },
            includeFollowingCommand ? first + "\n" + """{"command":"diag.ping"}""" : first);

        Assert.Equal(1, result.ExitCode);
        var continued = includeFollowingCommand && !stopOnError;
        AssertProgress(result.Stderr, includeFollowingCommand ? 2 : 1,
            continued ? [(2, "diag.ping", true)] : []);
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(continued ? 2 : 1, lines.Length);
        using var failed = JsonDocument.Parse(lines[0]);
        Assert.Equal(0, failed.RootElement.GetProperty("index").GetInt32());
        Assert.Equal("nonexistent.action", failed.RootElement.GetProperty("command").GetString());
        Assert.False(failed.RootElement.TryGetProperty("commandSummary", out _));
        Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(originalJson.RootElement.GetProperty("error").GetProperty("message").GetString(),
            failed.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.Equal(continued ? 1 : 0, requests.Count);
        if (continued)
        {
            Assert.Equal("diag.ping", requests[0].Command);
            using var next = JsonDocument.Parse(lines[1]);
            Assert.Equal(1, next.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("diag.ping", next.RootElement.GetProperty("command").GetString());
            Assert.True(next.RootElement.GetProperty("success").GetBoolean());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Batch_LaterValidationFailure_KeepsPriorResultsAndIndexes(bool stopOnError)
    {
        var requests = new List<ServiceRequest>();
        var result = await InProcessCliHelper.RunAsync(
            stopOnError ? ["batch", "--stop-on-error"] : ["batch"],
            request =>
            {
                requests.Add(request);
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            },
            """
            {"command":"diag.ping"}
            {"command":"range.get-values","args":{"sheetName":"Sales"}}
            {"command":"diag.ping"}
            """);

        Assert.Equal(1, result.ExitCode);
        AssertProgress(result.Stderr, 3, stopOnError
            ? [(1, "diag.ping", true)]
            : [(1, "diag.ping", true), (3, "diag.ping", true)]);
        Assert.Equal(stopOnError ? 1 : 2, requests.Count);
        Assert.All(requests, request => Assert.Equal("diag.ping", request.Command));
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(stopOnError ? 2 : 3, lines.Length);
        using var first = JsonDocument.Parse(lines[0]);
        Assert.Equal(0, first.RootElement.GetProperty("index").GetInt32());
        Assert.True(first.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("diag.ping", first.RootElement.GetProperty("command").GetString());
        using var failed = JsonDocument.Parse(lines[1]);
        Assert.Equal(1, failed.RootElement.GetProperty("index").GetInt32());
        Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("range.get-values", failed.RootElement.GetProperty("command").GetString());
        Assert.Contains("rangeAddress", failed.RootElement.GetProperty("error").GetProperty("message").GetString());
        if (!stopOnError)
        {
            using var last = JsonDocument.Parse(lines[2]);
            Assert.Equal(2, last.RootElement.GetProperty("index").GetInt32());
            Assert.True(last.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("diag.ping", last.RootElement.GetProperty("command").GetString());
        }
    }

    private static void AssertProgress(string stderr, int total,
        params (int Index, string Command, bool Success)[] items)
    {
        var expected = items.SelectMany(item => new[]
        {
            $"Batch {item.Index}/{total}: {item.Command} started.",
            $"Batch {item.Index}/{total}: {item.Command} {(item.Success ? "completed" : "failed")}."
        });
        Assert.Equal(expected, stderr.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries));
    }
}
