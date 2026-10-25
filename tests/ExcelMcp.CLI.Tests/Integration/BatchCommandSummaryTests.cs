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
public sealed class BatchCommandSummaryTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Batch_RepeatedCommands_HaveDistinctSummaries(bool arrayInput, bool fileInput)
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
            Assert.Empty(result.Stderr);
            var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            Assert.Equal(2, lines.Length);
            var summaries = new[]
            {
                "range.set-values: Sales!A2:B3 (2 rows x 2 columns)",
                "range.set-values: Forecast!B5:C5 (1 row x 2 columns)"
            };
            for (var index = 0; index < lines.Length; index++)
            {
                using var json = JsonDocument.Parse(lines[index]);
                var item = json.RootElement;
                Assert.Equal(index, item.GetProperty("index").GetInt32());
                Assert.Equal("range.set-values", item.GetProperty("command").GetString());
                Assert.Equal(summaries[index], item.GetProperty("commandSummary").GetString());
                Assert.True(item.GetProperty("success").GetBoolean());
                Assert.True(item.GetProperty("result").GetProperty("success").GetBoolean());
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
    public async Task Batch_Failures_KeepSummariesAndStopBehavior(string failure, bool stopOnError)
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
        Assert.Empty(result.Stderr);
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(stopOnError ? 1 : 2, lines.Length);
        using var json = JsonDocument.Parse(lines[0]);
        var item = json.RootElement;
        Assert.Equal(0, item.GetProperty("index").GetInt32());
        Assert.Equal("range.get-values", item.GetProperty("command").GetString());
        Assert.Equal(failure == "validation" ? "range.get-values: sheetName=Sales" : "range.get-values: Sales!A1",
            item.GetProperty("commandSummary").GetString());
        Assert.False(item.GetProperty("success").GetBoolean());
        Assert.False(string.IsNullOrWhiteSpace(item.GetProperty("errorMessage").GetString()));
        if (failure == "service")
        {
            Assert.Equal("Synthetic service failure", item.GetProperty("errorMessage").GetString());
            Assert.Equal("Synthetic", item.GetProperty("errorCategory").GetString());
            Assert.Equal("SyntheticException", item.GetProperty("exceptionType").GetString());
            Assert.Equal("0x800A03EC", item.GetProperty("hresult").GetString());
            Assert.Equal("Synthetic cause", item.GetProperty("innerError").GetString());
            Assert.True(item.GetProperty("partialState").GetProperty("changed").GetBoolean());
        }
        if (failure == "communication")
        {
            Assert.Equal("Communication error: Synthetic pipe failure", item.GetProperty("errorMessage").GetString());
        }
        if (failure == "negative-result")
        {
            Assert.Equal("Synthetic operation failure", item.GetProperty("errorMessage").GetString());
            Assert.False(item.GetProperty("result").GetProperty("success").GetBoolean());
        }
        Assert.Equal((failure == "validation" ? 0 : 1) + (stopOnError ? 0 : 1), requests.Count);
        if (!stopOnError)
        {
            using var next = JsonDocument.Parse(lines[1]);
            Assert.Equal(1, next.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("diag.ping", next.RootElement.GetProperty("commandSummary").GetString());
            Assert.True(next.RootElement.GetProperty("success").GetBoolean());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Batch_AllInvalid_ProducesSummariesWithoutConnecting(bool stopOnError)
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
        Assert.Equal("range.get-values: sheetName=Sales", first.RootElement.GetProperty("commandSummary").GetString());
        if (!stopOnError)
        {
            using var second = JsonDocument.Parse(lines[1]);
            Assert.Equal("range.get-values: sheetName=Forecast", second.RootElement.GetProperty("commandSummary").GetString());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Batch_LaterValidationFailure_KeepsPriorResultsAndSummary(bool stopOnError)
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
        Assert.Empty(result.Stderr);
        Assert.Equal(stopOnError ? 1 : 2, requests.Count);
        Assert.All(requests, request => Assert.Equal("diag.ping", request.Command));
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(stopOnError ? 2 : 3, lines.Length);
        using var first = JsonDocument.Parse(lines[0]);
        Assert.Equal(0, first.RootElement.GetProperty("index").GetInt32());
        Assert.True(first.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("diag.ping", first.RootElement.GetProperty("commandSummary").GetString());
        using var failed = JsonDocument.Parse(lines[1]);
        Assert.Equal(1, failed.RootElement.GetProperty("index").GetInt32());
        Assert.False(failed.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("range.get-values: sheetName=Sales", failed.RootElement.GetProperty("commandSummary").GetString());
        Assert.Contains("rangeAddress", failed.RootElement.GetProperty("errorMessage").GetString());
        if (!stopOnError)
        {
            using var last = JsonDocument.Parse(lines[2]);
            Assert.Equal(2, last.RootElement.GetProperty("index").GetInt32());
            Assert.True(last.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("diag.ping", last.RootElement.GetProperty("commandSummary").GetString());
        }
    }
}
