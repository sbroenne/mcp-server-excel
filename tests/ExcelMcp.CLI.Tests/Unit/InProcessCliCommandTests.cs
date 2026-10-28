using System.Text.Json;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "CommandRuntime")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
[Collection("Sequential")]
public sealed class InProcessCliCommandTests
{
    [Theory]
    [InlineData("powerquery", "list", "queries.name")]
    [InlineData("session", "list", "sessions.workbookName")]
    [InlineData("service", "status", "running")]
    public async Task RemovedFieldsOption_IsRejectedWithoutConnecting(string category, string action, string field)
    {
        var response = new ServiceResponse { Success = true, Result = """{"success":true,"queries":[],"sessions":[]}""" };
        var factory = new RecordingClientFactory(response);
        var daemon = new RecordingDaemonConnection(response,
            new DaemonConnectionPolicy.DaemonFailureState(DaemonConnectionPolicy.RunningState, true));
        var output = new StringWriter();
        string[] arguments = category == "powerquery"
            ? [category, action, "--session", "target", "--fields", field]
            : [category, action, "--fields", field];
        var exit = await Program.RunAsync(
            arguments, CreateRuntime(factory, output, new StringWriter(), daemonConnection: daemon));
        Assert.NotEqual(0, exit);
        Assert.Empty(factory.Requests);
        Assert.Empty(daemon.Requests);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("fields", result.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task SessionClose_PreservesActualOutcomeRatherThanInventingSavedState()
    {
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"message":"Session already closed.","warnings":["No save was performed."],"partialState":{"changed":false}}"""
        });
        var output = new StringWriter();
        var exit = await Program.RunAsync(
            ["session", "close", "--session", "target", "--save"],
            CreateRuntime(factory, output, new StringWriter()));
        Assert.Equal(0, exit);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal("Session already closed.", result.RootElement.GetProperty("data").GetProperty("message").GetString());
        Assert.Equal("No save was performed.", Assert.Single(result.RootElement.GetProperty("warnings").EnumerateArray()).GetString());
        Assert.False(result.RootElement.GetProperty("partialState").GetProperty("changed").GetBoolean());
        Assert.False(result.RootElement.GetProperty("data").TryGetProperty("saved", out _));
    }

    [Fact]
    public async Task SessionClose_DeliveredNegativeOutcomeDoesNotReportSaved()
    {
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":false,"errorMessage":"Workbook remains open","partialState":{"open":true}}"""
        });
        var output = new StringWriter();
        var exit = await Program.RunAsync(
            ["session", "close", "--session", "target", "--save"],
            CreateRuntime(factory, output, new StringWriter()));
        Assert.Equal(1, exit);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal("Workbook remains open", result.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.True(result.RootElement.GetProperty("partialState").GetProperty("open").GetBoolean());
        Assert.False(result.RootElement.TryGetProperty("data", out _));
    }

    [Fact]
    public async Task OutputFile_ImageKeepsCompleteBytesAndMetadata()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelcli-image-{Guid.NewGuid():N}.png");
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"imageBase64":"AQID","sheetName":"Data","width":1,"height":3,"mimeType":"image/png","message":"Truncated to the top-left portion."}"""
        });
        var output = new StringWriter();
        try
        {
            var exit = await Program.RunAsync(
                ["--quiet", "screenshot", "capture", "--session", "target", "--sheet-name", "Data",
                    "--range-address", "A1", "--output", path],
                CreateRuntime(factory, output, new StringWriter()));
            Assert.Equal(0, exit);
            Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(path));
            using var result = JsonDocument.Parse(output.ToString());
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(path, result.RootElement.GetProperty("data").GetProperty("outputPath").GetString());
            Assert.Equal("Truncated to the top-left portion.", result.RootElement.GetProperty("data").GetProperty("message").GetString());
            Assert.DoesNotContain("AQID", output.ToString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task OutputFile_WriteFailureRetainsCompletedOperation()
    {
        var path = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}", "result.json");
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"modifiedCells":3,"warnings":["Check retained state"],"partialState":{"changed":true}}"""
        });
        var output = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "range", "set-values", "--session", "target", "--sheet-name", "Data",
                "--range-address", "A1:C1", "--values", "[[1,2,3]]", "--output", path],
            CreateRuntime(factory, output, new StringWriter()));
        Assert.Equal(1, exit);
        Assert.Single(factory.Requests);
        using var result = JsonDocument.Parse(output.ToString());
        var root = result.RootElement;
        Assert.Equal("OutputWriteFailed", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal(3, root.GetProperty("data").GetProperty("modifiedCells").GetInt32());
        Assert.True(root.GetProperty("partialState").GetProperty("operationCompleted").GetBoolean());
        Assert.False(root.GetProperty("partialState").GetProperty("outputWriteCompleted").GetBoolean());
        Assert.True(root.GetProperty("partialState").GetProperty("operationPartialState").GetProperty("changed").GetBoolean());
        Assert.Equal("Check retained state", root.GetProperty("warnings")[0].GetString());
    }

    [Theory]
    [InlineData("\"0.00%\"")]
    [InlineData("null")]
    [InlineData("[false,0,null]")]
    public async Task OutputFile_WriteFailureKeepsOriginalDataShape(string data)
    {
        var path = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}", "result.json");
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = data });
        var output = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "chartconfig", "get-axis-number-format", "--session", "target",
                "--chart-name", "Revenue", "--axis", "Value", "--output", path],
            CreateRuntime(factory, output, new StringWriter()));
        Assert.Equal(1, exit);
        Assert.Single(factory.Requests);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal("OutputWriteFailed", result.RootElement.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal(data, result.RootElement.GetProperty("data").GetRawText());
        Assert.True(result.RootElement.GetProperty("partialState").GetProperty("operationCompleted").GetBoolean());
        Assert.False(result.RootElement.GetProperty("partialState").GetProperty("outputWriteCompleted").GetBoolean());
        Assert.False(File.Exists(path));
    }

    [Fact]
    public async Task NonRedirectedOperation_WritesOnlyJsonToStdout()
    {
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"message":"pong"}"""
        });
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = new CliCommandRuntime(factory, new StringReader(""), output, error, isOutputRedirected: false);
        Assert.Equal(0, await Program.RunAsync(["diag", "ping"], runtime));
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal("diag.ping", result.RootElement.GetProperty("command").GetString());
        Assert.Equal("pong", result.RootElement.GetProperty("data").GetProperty("message").GetString());
        Assert.Contains("Excel", error.ToString());
    }

    [Theory]
    [InlineData("powerquery", "list")]
    [InlineData("range", "get-values")]
    [InlineData("chartconfig", "get-axis-number-format")]
    [InlineData("session", "discover")]
    [InlineData("session", "list")]
    [InlineData("service", "status")]
    public async Task ReadHelp_OmitsRemovedFieldsOption(string category, string action)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        Assert.Equal(0, await Program.RunAsync(
            [category, action, "--help"], CreateRuntime(factory, output, new StringWriter())));
        Assert.Empty(factory.Requests);
        Assert.DoesNotContain("--fields", output.ToString());
    }

    [Fact]
    public async Task ManualRead_KeepsSessionIdentityAndRecoveryDetails()
    {
        var daemon = new RecordingDaemonConnection(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"count":1,"sessions":[{"sessionId":"target","workbookName":"Selected","filePath":"omitted.xlsx","canClose":false,"operationPending":true}]}"""
        }, new DaemonConnectionPolicy.DaemonFailureState(DaemonConnectionPolicy.RunningState, true));
        var output = new StringWriter();
        Assert.Equal(0, await Program.RunAsync(
            ["session", "list"],
            CreateRuntime(new RecordingClientFactory(), output, new StringWriter(), daemonConnection: daemon)));
        using var result = JsonDocument.Parse(output.ToString());
        var session = Assert.Single(result.RootElement.GetProperty("data").GetProperty("sessions").EnumerateArray());
        Assert.Equal("target", session.GetProperty("sessionId").GetString());
        Assert.Equal("Selected", session.GetProperty("workbookName").GetString());
        Assert.False(session.GetProperty("canClose").GetBoolean());
        Assert.True(session.GetProperty("operationPending").GetBoolean());
        Assert.Equal("omitted.xlsx", session.GetProperty("filePath").GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task QueryList_OmitsPreviewAndKeepsMetadata(bool batch)
    {
        const string payload = """{"success":true,"queries":[{"name":"Source","formulaPreview":"let x = 1 in x","loadMode":"connection-only"}]}""";
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = payload });
        var output = new StringWriter();
        var exit = await Program.RunAsync(
            batch ? ["batch", "--session", "target"] : ["--quiet", "powerquery", "list", "--session", "target"],
            CreateRuntime(factory, output, new StringWriter(), """{"command":"powerquery.list"}"""));
        Assert.Equal(0, exit);
        using var result = JsonDocument.Parse(output.ToString());
        var query = Assert.Single(result.RootElement.GetProperty("data").GetProperty("queries").EnumerateArray());
        Assert.Equal("Source", query.GetProperty("name").GetString());
        Assert.Equal("connection-only", query.GetProperty("loadMode").GetString());
        Assert.False(query.TryGetProperty("formulaPreview", out _));
        Assert.Null(Assert.Single(factory.Requests).Args);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Results_UseOneEnvelopeAndPreserveMeaningfulValues(bool batch)
    {
        const string payload = """{"success":true,"filePath":"C:\\fixture.xlsx","values":[[false,0,null]],"rowCount":1,"columnCount":3}""";
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = payload });
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = CreateRuntime(factory, output, error,
            """{"command":"range.get-values","sessionId":"target","args":{"sheetName":"Sheet1","rangeAddress":"A1:C1"}}""");
        var exit = await Program.RunAsync(
            batch ? ["--quiet", "batch"] :
                ["--quiet", "range", "get-values", "--session", "target", "--sheet-name", "Sheet1", "--range-address", "A1:C1"],
            runtime);
        Assert.Equal(0, exit);
        using var result = JsonDocument.Parse(output.ToString());
        var root = result.RootElement;
        Assert.Equal("range.get-values", root.GetProperty("command").GetString());
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("durationMs").GetInt64() >= 0);
        Assert.Equal(batch, root.TryGetProperty("index", out _));
        Assert.False(root.TryGetProperty("commandSummary", out _));
        Assert.False(root.TryGetProperty("result", out _));
        var data = root.GetProperty("data");
        Assert.False(data.TryGetProperty("success", out _));
        Assert.False(data.TryGetProperty("filePath", out _));
        Assert.Equal("[[false,0,null]]", data.GetProperty("values").GetRawText());
        Assert.Equal(1, data.GetProperty("rowCount").GetInt32());
        Assert.Equal(3, data.GetProperty("columnCount").GetInt32());
    }

    [Theory]
    [InlineData("enable-editing", "--allow-reopen")]
    [InlineData("enable-connections", "--allow-reopen")]
    [InlineData("enable-macros", "--allow-remembered-trust")]
    public async Task SessionSecurity_RejectsApprovalForADifferentAction(string action, string flag)
    {
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true });
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "session", action, "--session", "target", "--approved", flag],
            CreateRuntime(factory, output, error));
        Assert.NotEqual(0, exit);
        Assert.Empty(factory.Requests);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("only for", result.RootElement.GetProperty("error").GetProperty("message").GetString(), StringComparison.Ordinal);
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData("notices", false)]
    [InlineData("enable-editing", false)]
    [InlineData("enable-connections", false)]
    [InlineData("enable-macros", false)]
    [InlineData("enable-editing", true)]
    [InlineData("enable-connections", true)]
    [InlineData("enable-macros", true)]
    public async Task SessionSecurity_MapsExactActionAndExplicitApprovals(string action, bool approve)
    {
        const string result = """{"success":true,"security":{"changed":false}}""";
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var error = new StringWriter();
        var arguments = new List<string> { "--quiet", "session", action, "--session", "target" };
        if (approve)
            arguments.AddRange(["--approved", action == "enable-macros" ? "--allow-reopen" : "--allow-remembered-trust"]);
        var exit = await Program.RunAsync(arguments.ToArray(), CreateRuntime(factory, output, error));
        Assert.Equal(0, exit);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("session." + action, request.Command);
        Assert.Equal("target", request.SessionId);
        if (action == "notices")
            Assert.Null(request.Args);
        else
        {
            using var args = JsonDocument.Parse(request.Args!);
            Assert.Equal(approve, args.RootElement.GetProperty("approved").GetBoolean());
            Assert.Equal(approve, args.RootElement.GetProperty(
                action == "enable-macros" ? "allowReopen" : "allowRememberedTrust").GetBoolean());
            Assert.Equal(2, args.RootElement.EnumerateObject().Count());
        }
        AssertResult(result, output.ToString());
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task WorkbookAutoSave_MapsExplicitSettingAndReturnsStatus(bool enabled)
    {
        var result = JsonSerializer.Serialize(new { success = true, autoSaveOn = enabled });
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "workbook", "set-auto-save", "--session", "target",
                "--enabled", enabled ? "true" : "false"],
            CreateRuntime(factory, output, error));
        Assert.Equal(0, exit);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("workbook.set-auto-save", request.Command);
        Assert.Equal("target", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal(enabled, args.RootElement.GetProperty("enabled").GetBoolean());
        AssertResult(result, output.ToString());
        Assert.Empty(error.ToString());
    }

    [Fact]
    public async Task ConnectionAccountSettings_SetMapsTypedInputsWithoutEchoingAccount()
    {
        const string result = """{"success":true,"changed":true,"accountHintPresent":true,"interactiveLogin":"Always","identityMode":"CurrentUser"}""";
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "connection", "set-account-settings", "--session", "account-session",
                "--connection-name", "Sales", "--account-hint", "fixture-account",
                "--interactive-login", "Always", "--identity-mode", "CurrentUser"],
            CreateRuntime(factory, output, error));
        Assert.Equal(0, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("connection.set-account-settings", request.Command);
        Assert.Equal("account-session", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("Sales", args.RootElement.GetProperty("connectionName").GetString());
        Assert.Equal("fixture-account", args.RootElement.GetProperty("accountHint").GetString());
        Assert.Equal("Always", args.RootElement.GetProperty("interactiveLogin").GetString());
        Assert.Equal("CurrentUser", args.RootElement.GetProperty("identityMode").GetString());
        using var actual = JsonDocument.Parse(output.ToString());
        Assert.True(actual.RootElement.GetProperty("data").GetProperty("changed").GetBoolean());
        Assert.Equal("Always", actual.RootElement.GetProperty("data").GetProperty("interactiveLogin").GetString());
        Assert.DoesNotContain("fixture-account", output.ToString(), StringComparison.Ordinal);
        Assert.Empty(error.ToString());
    }

    [Fact]
    public async Task SessionDiscoverAndAttach_MapCommandsAndWorkbookIdentifier()
    {
        const string discoverResult = """{"success":true,"complete":false,"workbooks":[],"errors":[{"processId":1234,"message":"Window callback failed."}]}""";
        var discoverFactory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = discoverResult
        });
        var discoverOutput = new StringWriter();
        var discoverError = new StringWriter();
        var discoverExitCode = await Program.RunAsync(
            ["--quiet", "session", "discover"],
            CreateRuntime(discoverFactory, discoverOutput, discoverError));

        Assert.Equal(0, discoverExitCode);
        var discoverRequest = Assert.Single(discoverFactory.Requests);
        Assert.Equal("session.discover", discoverRequest.Command);
        Assert.Null(discoverRequest.SessionId);
        Assert.Null(discoverRequest.Args);
        AssertResult(discoverResult, discoverOutput.ToString());
        Assert.Empty(discoverError.ToString());

        const string attachResult = """{"success":true,"sessionId":"attached-session","name":"Book1","warning":"Excel AutoSave is off."}""";
        var attachFactory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            SessionId = "attached-session",
            Result = attachResult
        });
        var attachOutput = new StringWriter();
        var attachError = new StringWriter();
        var attachExitCode = await Program.RunAsync(
            ["--quiet", "session", "attach", "--workbook-id", "open-workbook-123"],
            CreateRuntime(attachFactory, attachOutput, attachError));

        Assert.Equal(0, attachExitCode);
        var attachRequest = Assert.Single(attachFactory.Requests);
        Assert.Equal("session.attach", attachRequest.Command);
        Assert.Null(attachRequest.SessionId);
        using var args = JsonDocument.Parse(attachRequest.Args!);
        Assert.Equal("open-workbook-123", args.RootElement.GetProperty("workbookId").GetString());
        AssertResult(attachResult, attachOutput.ToString());
        Assert.Empty(attachError.ToString());
    }

    [Theory]
    [InlineData("get-account-settings", """{"success":true,"accountHintPresent":true,"passwordPresent":false,"impersonationPresent":false,"identityMode":"Connection"}""")]
    [InlineData("clear-account-hint", """{"success":true,"changed":true,"accountHintPresent":false}""")]
    public async Task ConnectionAccountSettings_MapsExactConnectionAndPreservesResults(string action, string result)
    {
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "connection", action, "--session", "account-session", "--connection-name", "Sales"],
            CreateRuntime(factory, output, error));
        Assert.Equal(0, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("connection." + action, request.Command);
        Assert.Equal("account-session", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("Sales", args.RootElement.GetProperty("connectionName").GetString());
        Assert.Single(args.RootElement.EnumerateObject());
        AssertResult(result, output.ToString());
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SessionClose_NativeRefreshBusy_PreservesFailureAndExitCode(bool save)
    {
        const string message = "Cannot close: an Excel refresh is running. Changes remain in the open workbook.";
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = false,
            Command = "session.close",
            SessionId = "busy-session",
            ErrorCategory = "Busy",
            ErrorMessage = message
        });
        var output = new StringWriter();
        var error = new StringWriter();
        string[] arguments = save
            ? ["--quiet", "session", "close", "--session", "busy-session", "--save"]
            : ["--quiet", "session", "close", "--session", "busy-session"];
        var exitCode = await Program.RunAsync(arguments, CreateRuntime(factory, output, error));

        Assert.Equal(1, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("session.close", request.Command);
        Assert.Equal("busy-session", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal(save, args.RootElement.GetProperty("save").GetBoolean());
        using var result = JsonDocument.Parse(output.ToString());
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("Busy", result.RootElement.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal(message, result.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.False(result.RootElement.TryGetProperty("saved", out _));
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData("busy")]
    [InlineData("dialogOpen")]
    public async Task SessionList_NativeRefreshWithoutTrackedWork_PreservesCannotClose(string excelState)
    {
        var daemonConnection = new RecordingDaemonConnection(new ServiceResponse
        {
            Success = true,
            Result = $$"""{"success":true,"sessions":[{"sessionId":"busy-session","activeOperations":0,"operationPending":false,"canClose":false,"excelState":"{{excelState}}","blockingReason":"Check Excel before retrying."}]}"""
        }, new DaemonConnectionPolicy.DaemonFailureState(DaemonConnectionPolicy.RunningState, true));
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "list"],
            CreateRuntime(new RecordingClientFactory(), output, error, daemonConnection: daemonConnection));
        Assert.Equal(0, exitCode);
        Assert.Equal("session.list", Assert.Single(daemonConnection.Requests).Command);
        using var result = JsonDocument.Parse(output.ToString());
        var session = Assert.Single(result.RootElement.GetProperty("data").GetProperty("sessions").EnumerateArray());
        Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
        Assert.False(session.GetProperty("operationPending").GetBoolean());
        Assert.False(session.GetProperty("canClose").GetBoolean());
        Assert.Equal(excelState, session.GetProperty("excelState").GetString());
        Assert.Equal("Check Excel before retrying.", session.GetProperty("blockingReason").GetString());
        Assert.Empty(error.ToString());
    }

    [Fact]
    public async Task SessionList_PreservesGenericDialogAndPendingState()
    {
        var daemonConnection = new RecordingDaemonConnection(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"sessions":[{"sessionId":"pending-dialog","activeOperations":1,"operationPending":true,"canClose":false,"excelState":"dialogOpen","blockingReason":"dialogOpen","blockingMessage":"Check the Excel window and respond there.\nDialog text: The database is unavailable."}]}"""
        }, new DaemonConnectionPolicy.DaemonFailureState(DaemonConnectionPolicy.RunningState, true));
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "list"],
            CreateRuntime(new RecordingClientFactory(), output, error, daemonConnection: daemonConnection));

        Assert.Equal(0, exitCode);
        using var result = JsonDocument.Parse(output.ToString());
        var session = Assert.Single(result.RootElement.GetProperty("data").GetProperty("sessions").EnumerateArray());
        Assert.Equal("dialogOpen", session.GetProperty("excelState").GetString());
        Assert.Equal("dialogOpen", session.GetProperty("blockingReason").GetString());
        Assert.Equal("Check the Excel window and respond there.\nDialog text: The database is unavailable.", session.GetProperty("blockingMessage").GetString());
        Assert.True(session.GetProperty("operationPending").GetBoolean());
        Assert.Equal(1, session.GetProperty("activeOperations").GetInt32());
        Assert.False(session.GetProperty("canClose").GetBoolean());
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData("https://contoso-my.sharepoint.com/personal/test/Documents/Test.xlsx?web=1", true)]
    [InlineData("https://d.docs.live.net/123/Documents/Test.xlsx?download=1", false)]
    [InlineData("http://example.com:8080/Test.csv?download=1#data", false)]
    [InlineData(@"C:\Imports\Test.txt", false)]
    public async Task SessionOpen_ForwardsOriginalLocationAndVisibilityToSharedService(string url, bool show)
    {
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"sessionId":"remote-session"}"""
        });
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            show ? ["--quiet", "session", "open", url, "--show"] : ["--quiet", "session", "open", url],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("session.open", request.Command);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal(url, args.RootElement.GetProperty("filePath").GetString());
        Assert.Equal(show, args.RootElement.GetProperty("show").GetBoolean());
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal("remote-session", result.RootElement.GetProperty("data").GetProperty("sessionId").GetString());
        Assert.Empty(error.ToString());
    }

    [Fact]
    public async Task RunAsync_ReturnsBeforePendingCommandCompletesAndTracksFinalOutcome()
    {
        var response = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var returned = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var factory = new RecordingClientFactory { PendingResponse = response.Task };
        var output = new StringWriter();
        var error = new StringWriter();
        var telemetry = new List<(string Command, bool Succeeded, string? ErrorCategory)>();
        var runtime = CreateRuntime(
            factory, output, error,
            telemetryObserver: (command, _, succeeded, category, _) =>
                telemetry.Add((command, succeeded, category)));
        var execution = Task.Run(() =>
        {
            var invocation = Program.RunAsync(
                ["--quiet", "session", "open", @"C:\workbooks\book.xlsx"], runtime);
            returned.SetResult();
            return invocation;
        });

        bool returnedBeforeCompletion;
        try
        {
            await factory.RequestSent.Task.WaitAsync(TimeSpan.FromSeconds(10));
            returnedBeforeCompletion = await Task.WhenAny(returned.Task, Task.Delay(TimeSpan.FromSeconds(2))) == returned.Task;
            Assert.False(execution.IsCompleted);
            Assert.Empty(telemetry);
        }
        finally
        {
            response.SetResult(new ServiceResponse { Success = true, Result = """{"sessionId":"pending"}""" });
        }

        Assert.Equal(0, await execution.WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.True(returnedBeforeCompletion, "RunAsync blocked until the command completed.");
        Assert.Equal([("session.open", true, null)], telemetry);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal("pending", result.RootElement.GetProperty("data").GetProperty("sessionId").GetString());
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData(null, 0)]
    [InlineData("range", 0)]
    [InlineData("session", 0)]
    [InlineData("unknown", 1)]
    public void ListActions_WritesJsonToInjectedOutput(string? command, int expectedExitCode)
    {
        var output = new StringWriter();
        var error = new StringWriter();
        using var scope = CliCommandRuntime.Push(CreateRuntime(new RecordingClientFactory(), output, error));

        var exitCode = new ListActionsCommand().Execute(
            null!, new ListActionsCommand.Settings { CommandName = command }, CancellationToken.None);

        Assert.Equal(expectedExitCode, exitCode);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.Equal(expectedExitCode == 0, result.RootElement.GetProperty("success").GetBoolean());
        if (command == "range")
        {
            Assert.Equal("list-actions", result.RootElement.GetProperty("command").GetString());
            Assert.Equal("range", result.RootElement.GetProperty("data").GetProperty("command").GetString());
            Assert.Contains("get-values", result.RootElement.GetProperty("data").GetProperty("actions").EnumerateArray().Select(value => value.GetString()));
        }
        else if (command == "session")
        {
            var actions = result.RootElement.GetProperty("data").GetProperty("actions").EnumerateArray()
                .Select(value => value.GetString()).ToArray();
            Assert.Contains("discover", actions);
            Assert.Contains("attach", actions);
        }
        else if (command == null)
        {
            var actions = result.RootElement.GetProperty("data").GetProperty("commands").GetProperty("session")
                .EnumerateArray().Select(value => value.GetString()).ToArray();
            Assert.Contains("discover", actions);
            Assert.Contains("attach", actions);
            Assert.Contains("session discover", result.RootElement.GetProperty("data").GetProperty("workflow").GetString());
            Assert.Contains("session attach", result.RootElement.GetProperty("data").GetProperty("workflow").GetString());
        }
        else
        {
            Assert.Equal("Unknown command 'unknown'.", result.RootElement.GetProperty("error").GetProperty("message").GetString());
        }
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData("status")]
    [InlineData("stop")]
    public async Task StoppedService_WritesJsonToInjectedOutput(string action)
    {
        var previousPipe = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE");
        var output = new StringWriter();
        var error = new StringWriter();
        var factory = new RecordingClientFactory();
        try
        {
            Environment.SetEnvironmentVariable("EXCELMCP_CLI_PIPE", $"excelmcp-output-{Guid.NewGuid():N}");
            var exitCode = await Program.RunAsync(
                ["--quiet", "service", action], CreateRuntime(factory, output, error));

            Assert.Equal(0, exitCode);
            using var result = JsonDocument.Parse(output.ToString());
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            if (action == "status")
            {
                Assert.Equal("stopped", result.RootElement.GetProperty("data").GetProperty("daemonState").GetString());
                Assert.False(result.RootElement.GetProperty("data").GetProperty("running").GetBoolean());
                Assert.Equal(0, result.RootElement.GetProperty("data").GetProperty("processId").GetInt32());
            }
            else
            {
                Assert.Equal("Service not running.", result.RootElement.GetProperty("data").GetProperty("message").GetString());
            }
            Assert.Empty(factory.Requests);
            Assert.Empty(error.ToString());
        }
        finally
        {
            Environment.SetEnvironmentVariable("EXCELMCP_CLI_PIPE", previousPipe);
        }
    }

    [Fact]
    public async Task ServiceStop_NoService_ReturnsWithoutWaitingForConnectTimeout()
    {
        var previousPipe = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE");
        var output = new StringWriter();
        var factory = new RecordingClientFactory();
        try
        {
            Environment.SetEnvironmentVariable("EXCELMCP_CLI_PIPE", $"excelmcp-stop-{Guid.NewGuid():N}");
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var exitCode = await Program.RunAsync(
                ["--quiet", "service", "stop"], CreateRuntime(factory, output, new StringWriter()));
            stopwatch.Stop();

            Assert.Equal(0, exitCode);
            using var result = JsonDocument.Parse(output.ToString());
            Assert.Equal("Service not running.", result.RootElement.GetProperty("data").GetProperty("message").GetString());
            Assert.InRange(stopwatch.Elapsed, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }
        finally
        {
            Environment.SetEnvironmentVariable("EXCELMCP_CLI_PIPE", previousPipe);
        }
    }

    public static IEnumerable<object[]> GeneratedCommandNames =>
        _CliCategoryMetadata.ValidActionsByCommand.Keys
            .Order(StringComparer.Ordinal)
            .Select(command => new object[] { command });

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task BatchDiscovery_MainHelpAndInitialUsageExplainWhenToBatch(bool showHelp)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            showHelp ? ["--quiet", "--help"] : ["--quiet"],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exit);
        Assert.Equal(0, factory.ConnectionCount);
        var text = Regex.Replace(output.ToString() + error, @"\s+", " ");
        Assert.Contains("known sequence", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("batch --help", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task BatchDiscovery_BatchHelpExplainsInputSessionsAndFailures()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "batch", "--help"], CreateRuntime(factory, output, error));

        Assert.Equal(0, exit);
        Assert.Equal(0, factory.ConnectionCount);
        Assert.Empty(error.ToString());
        var text = output.ToString();
        foreach (var expected in new[]
                 {
                     "group.action", "command", "args", "sessionId", "sheetName", "rangeAddress",
                     "case-sensitive", "JSON array", "one command per line", "--input -",
                     "session.open/create/attach", "--stop-on-error", "default: continue",
                     "earlier changes", "index", "$LASTEXITCODE", "stderr", "expected",
                     "excelcli.exe batch", "range.set-values"
                 })
        {
            Assert.Contains(expected, text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task BatchDiscovery_DiagnosticExampleExecutesAsPrinted()
    {
        var help = new StringWriter();
        await Program.RunAsync(
            ["--quiet", "batch", "--help"],
            CreateRuntime(new RecordingClientFactory(), help, new StringWriter()));
        var example = Regex.Match(help.ToString(), @"(?s)@'\r?\n(.*?)\r?\n'@");
        Assert.True(example.Success, "Batch help must include a runnable PowerShell input example.");
        var input = example.Groups[1].Value;
        using var entries = JsonDocument.Parse(input);
        Assert.Equal(2, entries.RootElement.GetArrayLength());

        var factory = new RecordingClientFactory(
            new ServiceResponse { Success = true, Result = """{"message":"pong"}""" },
            new ServiceResponse { Success = true, Result = """{"message":"hello"}""" });
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "batch", "--stop-on-error"],
            CreateRuntime(factory, output, error, input));

        Assert.Equal(0, exit);
        Assert.Collection(factory.Requests,
            request => Assert.Equal("diag.ping", request.Command),
            request =>
            {
                Assert.Equal("diag.echo", request.Command);
                Assert.Equal("""{"message":"hello"}""", request.Args);
            });
        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        for (var index = 0; index < lines.Length; index++)
        {
            using var result = JsonDocument.Parse(lines[index]);
            Assert.Equal(index, result.RootElement.GetProperty("index").GetInt32());
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(index == 0 ? "pong" : "hello",
                result.RootElement.GetProperty("data").GetProperty("message").GetString());
        }
        Assert.Contains("started", error.ToString(), StringComparison.Ordinal);
        Assert.DoesNotContain("started", output.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(GeneratedCommandNames))]
    public async Task BatchDiscovery_ActionHelpMapsExactArgumentsWithoutConnecting(string command)
    {
        foreach (var (action, parameters) in _CliCategoryMetadata.ActionParametersByCommand[command])
        {
            var factory = new RecordingClientFactory();
            var output = new StringWriter();
            var error = new StringWriter();
            var exit = await Program.RunAsync(
                ["--quiet", command, action, "--help"], CreateRuntime(factory, output, error));

            Assert.Equal(0, exit);
            Assert.Equal(0, factory.ConnectionCount);
            Assert.Empty(error.ToString());
            var text = output.ToString();
            var section = text.IndexOf("Batch input", StringComparison.Ordinal);
            Assert.True(section >= 0, $"Missing batch help for {command}.{action}.");
            var batchHelp = text[section..];
            Assert.Contains($"\"command\":\"{command}.{action}\"", batchHelp, StringComparison.Ordinal);
            foreach (var parameter in parameters)
                Assert.Contains($"args.{parameter}", batchHelp, StringComparison.Ordinal);
            Assert.DoesNotContain("args.sessionId", batchHelp, StringComparison.Ordinal);
            Assert.DoesNotContain("args.outputPath", batchHelp, StringComparison.Ordinal);
            Assert.DoesNotContain("--output", batchHelp, StringComparison.Ordinal);
            var category = _CliCategoryMetadata.Categories.Single(item => item.CliCommandName == command);
            if (category.RequiresSession)
                Assert.Contains("--session", batchHelp, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task BatchDiscovery_RangeHelpIncludesValidNativeJsonExample()
    {
        var output = new StringWriter();
        await Program.RunAsync(
            ["--quiet", "range", "set-values", "--help"],
            CreateRuntime(new RecordingClientFactory(), output, new StringWriter()));

        var text = output.ToString();
        Assert.Contains("args.sheetName", text, StringComparison.Ordinal);
        Assert.Contains("args.rangeAddress", text, StringComparison.Ordinal);
        var example = text.Split('\n')
            .Last(line => line.StartsWith("{\"command\":\"range.set-values\"", StringComparison.Ordinal));
        using var entry = JsonDocument.Parse(example);
        var args = entry.RootElement.GetProperty("args");
        Assert.Equal("Sales", args.GetProperty("sheetName").GetString());
        Assert.Equal("A1:B2", args.GetProperty("rangeAddress").GetString());
        Assert.Equal(42, args.GetProperty("values")[1][1].GetInt32());
        ServiceRegistry.ValidateCommandArguments("range.set-values", args.GetRawText());
    }

    [Theory]
    [InlineData("diag ping", "Unknown command group", "Use 'diag.ping'")]
    [InlineData("DIAG PING", "Unknown command group", "Use 'diag.ping'")]
    [InlineData("diag\tping", "Unknown command group", "Use 'diag.ping'")]
    [InlineData("diag not-an-action", "Unknown command group", "group.action")]
    [InlineData("diag.not-an-action", "Unknown action", "group.action")]
    public async Task BatchDiscovery_CommandErrorsSuggestButNeverExecute(
        string command, string originalError, string hint)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var input = JsonSerializer.Serialize(new[] { new { command } });
        var exit = await Program.RunAsync(
            ["--quiet", "batch", "--stop-on-error"],
            CreateRuntime(factory, output, error, input));

        Assert.Equal(1, exit);
        Assert.Equal(0, factory.ConnectionCount);
        Assert.Empty(factory.Requests);
        Assert.Empty(error.ToString());
        using var result = JsonDocument.Parse(output.ToString());
        var item = result.RootElement;
        Assert.False(item.GetProperty("success").GetBoolean());
        Assert.Equal(0, item.GetProperty("index").GetInt32());
        Assert.Equal(command, item.GetProperty("command").GetString());
        Assert.True(item.GetProperty("durationMs").GetInt64() >= 0);
        var message = item.GetProperty("error").GetProperty("message").GetString();
        Assert.Contains(originalError, message, StringComparison.Ordinal);
        Assert.Contains(hint, message, StringComparison.Ordinal);
        Assert.Contains("batch --help", message, StringComparison.Ordinal);
        if (command.Contains("not-an-action", StringComparison.Ordinal))
            Assert.DoesNotContain("Use 'diag.not-an-action'", message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("[]", "No commands provided")]
    [InlineData("[", "BytePositionInLine")]
    [InlineData("""{"args":{}}""", "missing the 'command' field")]
    [InlineData("""{"command":"diag.ping","unexpected":true}""", "unexpected")]
    public async Task BatchDiscovery_InputErrorsKeepDetailsAndPointToHelp(string input, string originalError)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "batch"], CreateRuntime(factory, output, error, input));

        Assert.Equal(1, exit);
        Assert.Equal(0, factory.ConnectionCount);
        using var result = JsonDocument.Parse(output.ToString());
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        var failure = result.RootElement.GetProperty("error");
        var message = failure.GetProperty("message").GetString();
        Assert.Contains(originalError, message, StringComparison.Ordinal);
        Assert.Contains("batch --help", message + error.ToString(), StringComparison.Ordinal);
        if (input == "[" || input.Contains("unexpected", StringComparison.Ordinal))
            Assert.Equal(nameof(JsonException), failure.GetProperty("exceptionType").GetString());
    }

    [Fact]
    public async Task BatchDiscovery_WrongArgumentKeysPointToActionHelp()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exit = await Program.RunAsync(
            ["--quiet", "batch"],
            CreateRuntime(factory, output, error,
                """{"command":"range.set-values","args":{"sheet":"Sales","range":"A1","values":[[42]]}}"""));

        Assert.Equal(1, exit);
        Assert.Equal(0, factory.ConnectionCount);
        Assert.Empty(error.ToString());
        using var result = JsonDocument.Parse(output.ToString());
        var message = result.RootElement.GetProperty("error").GetProperty("message").GetString();
        Assert.Contains("Unknown or non-canonical", message, StringComparison.Ordinal);
        Assert.Contains("sheetName", message, StringComparison.Ordinal);
        Assert.Contains("rangeAddress", message, StringComparison.Ordinal);
        Assert.Contains("excelcli range set-values --help", message, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(GeneratedCommandNames))]
    public async Task CommandHelp_ListsEveryValidActionWithoutConnecting(string command)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", command, "--help"],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Equal(string.Empty, error.ToString());
        var help = Regex.Replace(output.ToString(), @"\s+", " ");
        var argumentSection = help.Split("OPTIONS:", StringSplitOptions.None)[0];
        Assert.Contains("Available actions:", argumentSection, StringComparison.Ordinal);
        foreach (var action in _CliCategoryMetadata.ValidActionsByCommand[command])
        {
            Assert.Matches($@"(?<![\w-]){Regex.Escape(action)}(?![\w-])", argumentSection);
        }
    }

    [Fact]
    public async Task ActionHelp_ListsOnlyOptionsForSelectedAction()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "slicer", "create-slicer", "--help"],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Empty(error.ToString());
        var help = output.ToString();
        Assert.Contains("Usage: excelcli slicer create-slicer", help, StringComparison.Ordinal);
        Assert.Contains("--pivot-table-name", help, StringComparison.Ordinal);
        Assert.Contains("--field-name", help, StringComparison.Ordinal);
        Assert.DoesNotContain("--selected-items", help, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ActionHelp_ShowsEveryGeneratedActionDescription()
    {
        foreach (var (command, actions) in _CliCategoryMetadata.ValidActionsByCommand)
        {
            foreach (var action in actions)
            {
                var factory = new RecordingClientFactory();
                var output = new StringWriter();
                var error = new StringWriter();

                var exitCode = await Program.RunAsync(
                    ["--quiet", command, action, "--help"],
                    CreateRuntime(factory, output, error));

                Assert.Equal(0, exitCode);
                Assert.Empty(factory.Requests);
                Assert.Empty(error.ToString());
                var description = _CliCategoryMetadata.ActionDescriptionsByCommand[command][action];
                Assert.False(string.IsNullOrWhiteSpace(description), $"{command} {action} has no description.");
                var beforeOptions = output.ToString().Split("Options:", StringSplitOptions.None)[0];
                Assert.Contains(description, beforeOptions, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    [Trait("Feature", "GeneratedContracts")]
    public async Task ActionHelp_ShowsCliAlternativesForSharedCommandReferences()
    {
        foreach (var (command, action, targetCommand, targetAction) in new[]
                 {
                     ("workbook", "create-cell-style", "rangeformat", "set-style"),
                     ("chartconfig", "add-series", "pivottablefield", "add-value-field"),
                     ("chartconfig", "remove-series", "pivottablefield", "remove-field")
                 })
        {
            var factory = new RecordingClientFactory();
            var output = new StringWriter();
            var error = new StringWriter();

            var exitCode = await Program.RunAsync(
                ["--quiet", command, action, "--help"],
                CreateRuntime(factory, output, error));

            Assert.Equal(0, exitCode);
            Assert.Empty(factory.Requests);
            Assert.Empty(error.ToString());
            var description = output.ToString().Split("Options:", StringSplitOptions.None)[0];
            Assert.Contains($"CLI: {targetCommand} {targetAction}", description, StringComparison.Ordinal);
            Assert.Contains(targetAction, _CliCategoryMetadata.ValidActionsByCommand[targetCommand]);
        }
    }

    [Fact]
    [Trait("Feature", "GeneratedContracts")]
    public async Task ActionHelp_ShowsCliFlagForSharedParameterReference()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "chartconfig", "set-style", "--help"],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Empty(error.ToString());
        var help = output.ToString().Split("Options:", StringSplitOptions.None);
        Assert.Contains("CLI: --style-id", help[0], StringComparison.Ordinal);
        Assert.Contains("--style-id", help[1], StringComparison.Ordinal);
    }

    [Fact]
    public async Task ActionHelp_ShowsActionDescriptionBeforeOptions()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "range", "set-values", "--help"],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Empty(error.ToString());
        var description = output.ToString().Split("Options:", StringSplitOptions.None)[0];
        Assert.Contains("Sets values in a range from 2D array or file.", description, StringComparison.Ordinal);
        Assert.Contains("restoration can fail without failing the write.", description, StringComparison.Ordinal);
        Assert.Contains("[[1,2],[3,4]]", description, StringComparison.Ordinal);
        Assert.Contains("\"2025-01-15\"", description, StringComparison.Ordinal);
        Assert.DoesNotContain("Gets values from a range", description, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("session", "list", "not a completed or failed operation")]
    [InlineData("session", "list", "activeOperations:0")]
    [InlineData("session", "list", "before repeating")]
    [InlineData("session", "notices", "complete:false")]
    [InlineData("session", "notices", "does not prove")]
    [InlineData("session", "enable-editing", "not confirmation prompts")]
    [InlineData("session", "enable-connections", "not confirmation prompts")]
    [InlineData("session", "enable-macros", "not confirmation prompts")]
    [InlineData("session", "enable-editing", "company policy")]
    [InlineData("session", "enable-connections", "company policy")]
    [InlineData("session", "enable-macros", "company policy")]
    public async Task CommandHelp_ExplainsResultLimitsAndApprovalWithoutConnecting(
        string command, string action, string guidance)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", command, action, "--help"], CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Empty(error.ToString());
        var help = Regex.Replace(output.ToString(), @"\s+", " ");
        Assert.Contains(guidance, help, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ActionHelp_PreservesJsonArrayExamples()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "range", "set-values", "--help"],
            CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Empty(error.ToString());
        Assert.Contains("[[1,2,3],[4,5,6]]", output.ToString(), StringComparison.Ordinal);
        Assert.DoesNotContain("[[[[1,2,3]]", output.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("session", "open", "--show")]
    [InlineData("session", "attach", "--workbook-id")]
    [InlineData("session", "close", "--save")]
    [InlineData("service", null, "stop")]
    [InlineData("range", null, "--sheet-name")]
    [InlineData("range", null, "--range-address")]
    [InlineData("range", null, "--values")]
    public async Task CommandHelp_ExposesLifecycleCommandsAndNativeOptionsWithoutConnecting(
        string command, string? branch, string option)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        string[] arguments = branch is null
            ? ["--quiet", command, "--help"]
            : ["--quiet", command, branch, "--help"];

        var exitCode = await Program.RunAsync(arguments, CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Equal(string.Empty, error.ToString());
        Assert.Contains(option, output.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("--axis", typeof(ChartAxisType))]
    [InlineData("--legend-position", typeof(LegendPosition))]
    [InlineData("--marker-style", typeof(MarkerStyle))]
    [InlineData("--plot-by", typeof(ChartPlotBy))]
    [InlineData("--display-blanks-as", typeof(ChartDisplayBlanksAs))]
    [InlineData("--area", typeof(ChartAreaTarget))]
    public async Task ChartHelp_AdvertisesAcceptedEnumNamesWithoutConnecting(string option, Type enumType)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "chartconfig", "--help"], CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Equal(string.Empty, error.ToString());
        Assert.Contains(option, output.ToString(), StringComparison.Ordinal);
        var help = Regex.Replace(output.ToString(), @"\s+", " ");
        foreach (var name in Enum.GetNames(enumType))
        {
            Assert.Matches($@"\b{Regex.Escape(name)}\b", help);
        }
    }

    [Theory]
    [InlineData("chartconfig", "--show-percentage", "Excel can reject", "no visual effect")]
    [InlineData("datamodel", "--format-type", "not substituted", "safe fallback")]
    public async Task CommandHelp_ExplainsAuditedFailureWithoutConnecting(
        string command, string option, string expected, string misleading)
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", command, "--help"], CreateRuntime(factory, output, error));
        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Equal(string.Empty, error.ToString());
        var help = Regex.Replace(output.ToString(), @"\s+", " ");
        Assert.Contains(option, help, StringComparison.Ordinal);
        Assert.Contains(expected, help, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(misleading, help, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ChartHelp_DoesNotAdvertiseReadOnlyLabelPosition()
    {
        var factory = new RecordingClientFactory();
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "chartconfig", "--help"], CreateRuntime(factory, output, error));
        Assert.Equal(0, exitCode);
        Assert.Empty(factory.Requests);
        Assert.Equal(string.Empty, error.ToString());
        var help = Regex.Replace(output.ToString(), @"\s+", " ");
        Assert.Contains("--label-position", help);
        foreach (var name in new[] { "BestFit", "Center", "Above", "Below", "Left", "Right", "InsideBase", "InsideEnd", "OutsideEnd" })
        {
            Assert.Matches($@"\b{Regex.Escape(name)}\b", help);
        }
        Assert.DoesNotContain("Mixed", help, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SessionOpen_ServiceFailure_PreservesRequestExitAndErrorEnvelope()
    {
        var factory = new RecordingClientFactory(
            new ServiceResponse
            {
                Success = false,
                Command = "session.open",
                ErrorCategory = "ComInterop",
                ErrorMessage = "Workbook is already open; close the file and retry with exclusive access.",
                ExceptionType = nameof(IOException)
            });
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = CreateRuntime(factory, output, error);
        const string workbookPath = @"C:\workbooks\locked.xlsx";

        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "open", workbookPath],
            runtime);

        Assert.Equal(1, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("session.open", request.Command);
        Assert.Null(request.SessionId);
        Assert.Equal(
            """{"filePath":"C:\\workbooks\\locked.xlsx","show":false}""",
            request.Args);

        using var envelope = JsonDocument.Parse(output.ToString());
        var root = envelope.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.False(root.TryGetProperty("isError", out _));
        Assert.Equal(
            "Workbook is already open; close the file and retry with exclusive access.",
            root.GetProperty("error").GetProperty("message").GetString());
        Assert.False(root.TryGetProperty("errorMessage", out _));
        Assert.Equal("ComInterop", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal(nameof(IOException), root.GetProperty("error").GetProperty("exceptionType").GetString());
        Assert.Equal("session.open", root.GetProperty("command").GetString());
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task Batch_UsesRealParserAndSendsEachValidatedRequest()
    {
        var factory = new RecordingClientFactory(
            new ServiceResponse { Success = true, Result = """{"message":"first"}""" },
            new ServiceResponse { Success = true, Result = """{"message":"second"}""" });
        var output = new StringWriter();
        var runtime = CreateRuntime(
            factory,
            output,
            new StringWriter(),
            """
            {"command":"diag.echo","args":{"message":"first"}}
            {"command":"diag.echo","args":{"message":"second"}}
            """);

        var exitCode = await Program.RunAsync(
            ["--quiet", "batch"],
            runtime);

        Assert.True(
            exitCode == 0,
            $"Exit: {exitCode}; output: {output}; requests: {factory.Requests.Count}");
        Assert.Collection(
            factory.Requests,
            request =>
            {
                Assert.Equal("diag.echo", request.Command);
                Assert.Null(request.SessionId);
                Assert.Equal("""{"message":"first"}""", request.Args);
                Assert.Equal("cli-batch", request.Source);
            },
            request =>
            {
                Assert.Equal("diag.echo", request.Command);
                Assert.Null(request.SessionId);
                Assert.Equal("""{"message":"second"}""", request.Args);
                Assert.Equal("cli-batch", request.Source);
            });

        var lines = output.ToString()
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        Assert.All(lines, line =>
        {
            using var result = JsonDocument.Parse(line);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        });
    }

    [Theory]
    [InlineData("--input")]
    [InlineData("-i")]
    public async Task Batch_SeparateStdinSentinel_UsesInjectedInput(string inputOption)
    {
        var factory = new RecordingClientFactory(
            new ServiceResponse { Success = true, Result = """{"message":"stdin"}""" });
        var output = new StringWriter();
        var runtime = CreateRuntime(
            factory,
            output,
            new StringWriter(),
            """{"command":"diag.echo","args":{"message":"stdin"}}""");

        var exitCode = await Program.RunAsync(
            ["--quiet", "batch", inputOption, "-"],
            runtime);

        Assert.Equal(0, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal("diag.echo", request.Command);
        Assert.Equal("""{"message":"stdin"}""", request.Args);
    }

    [Fact]
    public async Task NestedCollection_SeparateStdinSentinel_DispatchesPipedJson()
    {
        var factory = new RecordingClientFactory(
            new ServiceResponse { Success = true, Result = """{"success":true}""" });
        var output = new StringWriter();
        var originalInput = Console.In;
        Console.SetIn(new StringReader("""[["Release","Version"],["excelcli","2.0.13"]]"""));
        try
        {
            var exitCode = await Program.RunAsync(
                [
                    "--quiet",
                    "range",
                    "set-values",
                    "--session",
                    "session-1",
                    "--sheet-name",
                    "Data",
                    "--range-address",
                    "A1:B2",
                    "--values",
                    "-"
                ],
                CreateRuntime(factory, output, new StringWriter()));

            Assert.Equal(0, exitCode);
        }
        finally
        {
            Console.SetIn(originalInput);
        }

        var request = Assert.Single(factory.Requests);
        Assert.Equal("range.set-values", request.Command);
        Assert.Equal("session-1", request.SessionId);
        Assert.Equal(
            """{"sheetName":"Data","rangeAddress":"A1:B2","values":[["Release","Version"],["excelcli","2.0.13"]]}""",
            request.Args);
    }

    [Fact]
    public void StandaloneDashForUnrelatedOption_IsNotRewritten()
    {
        var args = new[] { "service", "run", "--pipe-name", "-" };

        var normalized = Program.NormalizeStandaloneDashOptionValues(args);

        Assert.Equal(args, normalized);
    }

    [Theory]
    [InlineData("open", false, "session.open")]
    [InlineData("open", true, "session.open")]
    [InlineData("create", false, "session.create")]
    [InlineData("create", true, "session.create")]
    public async Task SessionCommand_MapsShowFlag(
        string action,
        bool show,
        string expectedCommand)
    {
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"sessionId":"session-1"}"""
        });
        var output = new StringWriter();
        var runtime = CreateRuntime(factory, output, new StringWriter());
        var arguments = new List<string>
        {
            "--quiet",
            "session",
            action,
            @"C:\workbooks\book.xlsx"
        };
        if (show)
        {
            arguments.Add("--show");
        }
        var exitCode = await Program.RunAsync(arguments.ToArray(), runtime);

        Assert.Equal(0, exitCode);
        var request = Assert.Single(factory.Requests);
        Assert.Equal(expectedCommand, request.Command);
        Assert.Null(request.SessionId);
        Assert.Equal(
            $$"""{"filePath":"C:\\workbooks\\book.xlsx","show":{{show.ToString().ToLowerInvariant()}}}""",
            request.Args);
    }

    [Fact]
    public async Task Help_WritesParserOutputToInjectedOutput()
    {
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "--help"],
            CreateRuntime(new RecordingClientFactory(), output, error));

        Assert.Equal(0, exitCode);
        Assert.Contains("USAGE:", output.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session", output.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task Version_WritesFriendlyOutputToInjectedStreams()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = CreateRuntime(
            new RecordingClientFactory(),
            output,
            error,
            latestVersionProvider: static () => Task.FromResult<string?>(null));

        var exitCode = await Program.RunAsync(["--version"], runtime);

        Assert.Equal(0, exitCode);
        Assert.Contains("Could not check for updates", output.ToString(), StringComparison.Ordinal);
        Assert.Contains("Current version:", output.ToString(), StringComparison.Ordinal);
        var banner = error.ToString();
        Assert.False(string.IsNullOrWhiteSpace(banner));
        Assert.DoesNotContain(banner, output.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("stopped", false)]
    [InlineData("unresponsive", true)]
    public async Task SessionList_UsesControlledDaemonState(
        string expectedState,
        bool running)
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var daemonConnection = new RecordingDaemonConnection(
            new ServiceResponse
            {
                Success = false,
                ErrorCategory = "ServiceUnavailable",
                ErrorMessage = "controlled transport failure"
            },
            new DaemonConnectionPolicy.DaemonFailureState(expectedState, running));
        var runtime = CreateRuntime(
            new RecordingClientFactory(),
            output,
            error,
            daemonConnection: daemonConnection);

        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "list"],
            runtime);

        Assert.Equal(expectedState == "stopped" ? 0 : 1, exitCode);
        Assert.Equal("session.list", Assert.Single(daemonConnection.Requests).Command);
        using var envelope = JsonDocument.Parse(output.ToString());
        var data = envelope.RootElement.GetProperty("data");
        Assert.Equal(expectedState, data.GetProperty("daemonState").GetString());
        if (running)
        {
            Assert.True(data.GetProperty("running").GetBoolean());
            Assert.Equal(
                "controlled transport failure",
                envelope.RootElement.GetProperty("error").GetProperty("message").GetString());
        }
        else
        {
            Assert.Equal(0, data.GetProperty("count").GetInt32());
            Assert.Empty(data.GetProperty("sessions").EnumerateArray());
            Assert.False(data.TryGetProperty("running", out _));
        }
    }

    [Fact]
    public async Task SessionListCommand_StoppedService_ReturnsConfirmedEmptyList()
    {
        var output = new StringWriter();
        var daemonConnection = new RecordingDaemonConnection(
            new ServiceResponse
            {
                Success = false,
                ErrorCategory = "ServiceUnavailable",
                ErrorMessage = "controlled transport failure"
            },
            new DaemonConnectionPolicy.DaemonFailureState("stopped", false));
        using var scope = CliCommandRuntime.Push(CreateRuntime(
            new RecordingClientFactory(),
            output,
            new StringWriter(),
            daemonConnection: daemonConnection));

        var exitCode = await new SessionListCommand().ExecuteAsync(null!, CancellationToken.None);

        Assert.Equal(0, exitCode);
        Assert.Equal("session.list", Assert.Single(daemonConnection.Requests).Command);
        using var json = JsonDocument.Parse(output.ToString());
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("stopped", json.RootElement.GetProperty("data").GetProperty("daemonState").GetString());
        Assert.Equal(0, json.RootElement.GetProperty("data").GetProperty("count").GetInt32());
        Assert.Empty(json.RootElement.GetProperty("data").GetProperty("sessions").EnumerateArray());
        Assert.False(json.RootElement.TryGetProperty("errorMessage", out _));
    }

    [Fact]
    public async Task ServiceCommand_ResultReportsSuccessFalse_PrintsResultAndExitsNonzero()
    {
        const string result = """{"success":false,"errorMessage":"1 of 2 queries failed to refresh: 'Broken'.","refreshedQueries":["Good"],"skippedQueries":[],"failedQueries":[{"queryName":"Broken","errorCategory":"Expression","errorMessage":"[Expression.Error] bad","exceptionType":"COMException","hresult":"0x800A03EC"}]}""";
        var telemetry = new List<(string Command, bool Succeeded, string? ErrorCategory, bool ExpectedNegative)>();
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "powerquery", "refresh-all", "--session", "session-1"],
            CreateRuntime(factory, output, error,
                telemetryObserver: (command, _, succeeded, errorCategory, expectedNegative) =>
                    telemetry.Add((command, succeeded, errorCategory, expectedNegative))));

        Assert.Equal(1, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal("powerquery.refresh-all", Assert.Single(factory.Requests).Command);
        AssertResult(result, output.ToString());
        var tracked = Assert.Single(telemetry);
        Assert.False(tracked.Succeeded);
        Assert.False(tracked.ExpectedNegative);
    }

    [Fact]
    public async Task ServiceCommand_ResultReportsSuccessFalseWithOutputPath_DoesNotReportWrittenFile()
    {
        const string result = """{"success":false,"errorMessage":"Range is empty."}""";
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var outputPath = Path.Combine(Path.GetTempPath(), $"excelcli-negative-{Guid.NewGuid():N}.json");

        try
        {
            var exitCode = await Program.RunAsync(
                ["--quiet", "powerquery", "refresh-all", "--session", "session-1", "--output", outputPath],
                CreateRuntime(factory, output, new StringWriter()));

            Assert.Equal(1, exitCode);
            Assert.False(File.Exists(outputPath));
            AssertResult(result, output.ToString());
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public async Task Batch_ResultReportsSuccessFalse_MarksLineFailedAndHonorsStopOnError()
    {
        const string negative = """{"success":false,"errorMessage":"1 of 1 queries failed to refresh: 'Broken'.","errorCategory":"Expression"}""";
        var factory = new RecordingClientFactory(
            new ServiceResponse { Success = true, Result = negative },
            new ServiceResponse { Success = true, Result = """{"success":true}""" });
        var output = new StringWriter();
        var runtime = CreateRuntime(
            factory,
            output,
            new StringWriter(),
            """
            {"command":"powerquery.refresh-all","sessionId":"session-1"}
            {"command":"diag.echo","args":{"message":"after"}}
            """);

        var exitCode = await Program.RunAsync(["--quiet", "batch", "--stop-on-error"], runtime);

        Assert.Equal(1, exitCode);
        Assert.Equal("powerquery.refresh-all", Assert.Single(factory.Requests).Command);
        using var line = JsonDocument.Parse(Assert.Single(
            output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)));
        Assert.False(line.RootElement.GetProperty("success").GetBoolean());
        Assert.False(line.RootElement.TryGetProperty("result", out _));
        Assert.Equal(
            "1 of 1 queries failed to refresh: 'Broken'.",
            line.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task Batch_FileTestCanOpenFalse_MarksLineFailedAndContinues()
    {
        var factory = new RecordingClientFactory(
            new ServiceResponse { Success = true, Result = """{"success":false,"canOpen":false}""" },
            new ServiceResponse { Success = true, Result = """{"success":true}""" });
        var output = new StringWriter();
        var runtime = CreateRuntime(
            factory,
            output,
            new StringWriter(),
            """
            {"command":"session.test","args":{"filePath":"C:\\workbooks\\missing.xlsx"}}
            {"command":"diag.echo","args":{"message":"after"}}
            """);

        var exitCode = await Program.RunAsync(["--quiet", "batch"], runtime);

        Assert.Equal(1, exitCode);
        Assert.Equal(2, factory.Requests.Count);
        var lines = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        using (var first = JsonDocument.Parse(lines[0]))
        {
            Assert.False(first.RootElement.GetProperty("success").GetBoolean());
            Assert.False(first.RootElement.GetProperty("data").GetProperty("canOpen").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(first.RootElement.GetProperty("error").GetProperty("message").GetString()));
        }
        using (var second = JsonDocument.Parse(lines[1]))
        {
            Assert.True(second.RootElement.GetProperty("success").GetBoolean());
        }
    }

    [Fact]
    public async Task SessionTest_CanOpenFalse_TracksExpectedNegative()
    {
        var telemetry = new List<(string Command, bool Succeeded, string? ErrorCategory, bool ExpectedNegative)>();
        var runtime = CreateRuntime(
            new RecordingClientFactory(new ServiceResponse
            {
                Success = true,
                Result = """{"canOpen":false}"""
            }),
            new StringWriter(),
            new StringWriter(),
            telemetryObserver: (command, _, succeeded, errorCategory, expectedNegative) =>
                telemetry.Add((command, succeeded, errorCategory, expectedNegative)));

        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "test", @"C:\workbooks\missing.xlsx"],
            runtime);

        Assert.Equal(1, exitCode);
        Assert.Equal(
            [("session.test", true, null, true)],
            telemetry);
    }

    [Fact]
    public async Task SessionTest_MalformedSuccessResponse_TracksInvalidResponse()
    {
        var telemetry = new List<(string Command, bool Succeeded, string? ErrorCategory)>();
        var runtime = CreateRuntime(
            new RecordingClientFactory(new ServiceResponse
            {
                Success = true,
                Result = "{"
            }),
            new StringWriter(),
            new StringWriter(),
            telemetryObserver: (command, _, succeeded, errorCategory, _) =>
                telemetry.Add((command, succeeded, errorCategory)));

        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "test", @"C:\workbooks\book.xlsx"],
            runtime);

        Assert.NotEqual(0, exitCode);
        Assert.Equal(
            [("session.test", false, "InvalidResponse")],
            telemetry);
    }

    [Fact]
    public async Task SessionList_NullSuccessResponse_TracksInvalidResponse()
    {
        var telemetry = new List<(string Command, bool Succeeded, string? ErrorCategory)>();
        var daemonConnection = new RecordingDaemonConnection(
            new ServiceResponse { Success = true },
            new DaemonConnectionPolicy.DaemonFailureState("running", true));
        var runtime = CreateRuntime(
            new RecordingClientFactory(),
            new StringWriter(),
            new StringWriter(),
            daemonConnection: daemonConnection,
            telemetryObserver: (command, _, succeeded, errorCategory, _) =>
                telemetry.Add((command, succeeded, errorCategory)));

        var exitCode = await Program.RunAsync(
            ["--quiet", "session", "list"],
            runtime);

        Assert.Equal(1, exitCode);
        Assert.Equal(
            [("session.list", false, "InvalidResponse")],
            telemetry);
    }

    [Theory]
    [InlineData(null, "[\"2026 Q2\"]")]
    [InlineData(false, "[\"2026 Q1\"]")]
    [InlineData(true, "[]")]
    public async Task SlicerSelection_MapsNamesSelectionAndClearFirst(bool? clearFirst, string selectedItems)
    {
        const string result = """{"success":true,"availableItems":["2026 Q1","2026 Q2"],"selectedItems":["2026 Q2"],"connectedPivotTables":["RevenuePivot"]}""";
        var factory = new RecordingClientFactory(new ServiceResponse { Success = true, Result = result });
        var output = new StringWriter();
        var error = new StringWriter();
        var arguments = new List<string>
        {
            "--quiet", "slicer", "set-slicer-selection", "--session", "session-1",
            "--slicer-name", "QuarterSlicer", "--selected-items", selectedItems
        };
        if (clearFirst.HasValue)
            arguments.AddRange(["--clear-first", clearFirst.Value ? "true" : "false"]);

        var exitCode = await Program.RunAsync(arguments.ToArray(), CreateRuntime(factory, output, error));

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        var request = Assert.Single(factory.Requests);
        Assert.Equal("slicer.set-slicer-selection", request.Command);
        Assert.Equal("session-1", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("QuarterSlicer", args.RootElement.GetProperty("slicerName").GetString());
        Assert.Equal(selectedItems, args.RootElement.GetProperty("selectedItems").GetRawText());
        if (clearFirst.HasValue)
            Assert.Equal(clearFirst.Value, args.RootElement.GetProperty("clearFirst").GetBoolean());
        else
            Assert.False(args.RootElement.TryGetProperty("clearFirst", out _));
        using var returned = JsonDocument.Parse(output.ToString());
        Assert.True(returned.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("2026 Q2", returned.RootElement.GetProperty("data").GetProperty("selectedItems")[0].GetString());
        Assert.Equal("RevenuePivot", returned.RootElement.GetProperty("data").GetProperty("connectedPivotTables")[0].GetString());
    }

    [Fact]
    public async Task SlicerSelection_InvalidItemReturnsErrorAndNonzeroExit()
    {
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = false,
            Command = "slicer.set-slicer-selection",
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(ArgumentException),
            ErrorMessage = "Slicer item 'missing' was not found."
        });
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = await Program.RunAsync(
            ["--quiet", "slicer", "set-slicer-selection", "--session", "session-1",
                "--slicer-name", "QuarterSlicer", "--selected-items", "[\"missing\"]"],
            CreateRuntime(factory, output, error));

        Assert.Equal(1, exitCode);
        Assert.Equal("slicer.set-slicer-selection", Assert.Single(factory.Requests).Command);
        using var returned = JsonDocument.Parse(output.ToString());
        Assert.False(returned.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", returned.RootElement.GetProperty("error").GetProperty("category").GetString());
        Assert.Contains("missing", returned.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.Empty(error.ToString());
    }

    [Fact]
    public async Task SlicerSelection_AmbiguousErrorPreservesCandidateJsonAndRetryName()
    {
        string[] candidates =
        [
            "[Calendar].[Quarter].&[2025]&[Q\"1\\North]]]",
            "[Calendar].[Quarter].&[2026]&[Q\"1\\South]]]"
        ];
        const string marker = "matching MDX unique names: ";
        string message = new ArgumentException(
            $"Slicer caption 'Q1' is ambiguous. Retry with one of these {marker}{JsonSerializer.Serialize(candidates)}",
            "requested").Message;
        var factory = new RecordingClientFactory(new ServiceResponse
        {
            Success = false,
            Command = "slicer.set-slicer-selection",
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(ArgumentException),
            ErrorMessage = message
        }, new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"selectedItems":["Q1"]}"""
        });
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = await Program.RunAsync(
            ["--quiet", "slicer", "set-slicer-selection", "--session", "session-1",
                "--slicer-name", "QuarterSlicer", "--selected-items", "[\"Q1\"]"],
            CreateRuntime(factory, output, error));
        Assert.Equal(1, exitCode);
        Assert.Empty(error.ToString());
        using var returned = JsonDocument.Parse(output.ToString());
        Assert.False(returned.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", returned.RootElement.GetProperty("error").GetProperty("category").GetString());
        string returnedMessage = returned.RootElement.GetProperty("error").GetProperty("message").GetString()!;
        Assert.Equal(message, returnedMessage);
        var returnedCandidates = JsonSerializer.Deserialize<string[]>(returnedMessage[
            (returnedMessage.IndexOf(marker, StringComparison.Ordinal) + marker.Length)..(returnedMessage.LastIndexOf(']') + 1)])!;
        Assert.Equal(candidates, returnedCandidates);

        output.GetStringBuilder().Clear();
        exitCode = await Program.RunAsync(
            ["--quiet", "slicer", "set-slicer-selection", "--session", "session-1",
                "--slicer-name", "QuarterSlicer", "--selected-items", JsonSerializer.Serialize(new[] { returnedCandidates[1] })],
            CreateRuntime(factory, output, error));
        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(2, factory.Requests.Count);
        var retry = factory.Requests[1];
        Assert.Equal("slicer.set-slicer-selection", retry.Command);
        Assert.Equal("session-1", retry.SessionId);
        using var args = JsonDocument.Parse(retry.Args!);
        Assert.Equal("QuarterSlicer", args.RootElement.GetProperty("slicerName").GetString());
        Assert.Equal(candidates[1], Assert.Single(args.RootElement.GetProperty("selectedItems").EnumerateArray()).GetString());
        using var success = JsonDocument.Parse(output.ToString());
        Assert.True(success.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Batch_ReportsCurrentCommandAndDeliversEarlierResultWhileWaiting()
    {
        var waiting = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var calls = 0;
        var factory = new RecordingClientFactory
        {
            Responder = _ =>
            {
                if (Interlocked.Increment(ref calls) == 1)
                    return Task.FromResult(new ServiceResponse { Success = true, Result = """{"message":"first"}""" });
                waiting.SetResult();
                return release.Task;
            }
        };
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = CreateRuntime(factory, output, error, """
            {"command":"diag.echo","args":{"message":"first"}}
            {"command":"diag.echo","args":{"message":"second"}}
            """);
        var run = Program.RunAsync(["--quiet", "batch"], runtime);
        try
        {
            await waiting.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.False(run.IsCompleted);
            using var first = JsonDocument.Parse(Assert.Single(
                output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)));
            Assert.Equal(0, first.RootElement.GetProperty("index").GetInt32());
            Assert.Equal("first", first.RootElement.GetProperty("data").GetProperty("message").GetString());
            Assert.Contains("Batch 2/2: diag.echo", error.ToString());
        }
        finally
        {
            release.TrySetResult(new ServiceResponse { Success = true, Result = """{"message":"second"}""" });
            Assert.Equal(0, await run.WaitAsync(TimeSpan.FromSeconds(5)));
        }
        Assert.Equal(2, output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).Length);
    }

    private static void AssertResult(string serviceJson, string cliJson)
    {
        using var service = JsonDocument.Parse(serviceJson);
        using var cli = JsonDocument.Parse(cliJson);
        var root = cli.RootElement;
        Assert.Equal(service.RootElement.GetProperty("success").GetBoolean(), root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("durationMs").GetInt64() >= 0);
        foreach (var property in service.RootElement.EnumerateObject())
        {
            if (property.Name == "success")
                continue;
            if (property.Name is "errorMessage" or "errorCategory")
                Assert.Equal(property.Value.GetString(),
                    root.GetProperty("error").GetProperty(property.Name == "errorMessage" ? "message" : "category").GetString());
            else if (property.Name == "warning")
                Assert.Equal(property.Value.GetString(), Assert.Single(root.GetProperty("warnings").EnumerateArray()).GetString());
            else
                Assert.Equal(property.Value.GetRawText(), root.GetProperty("data").GetProperty(property.Name).GetRawText());
        }
    }

    private static CliCommandRuntime CreateRuntime(
        RecordingClientFactory factory,
        StringWriter output,
        StringWriter error,
        string input = "",
        ICliDaemonConnection? daemonConnection = null,
        Func<Task<string?>>? latestVersionProvider = null,
        Action<string, long, bool, string?, bool>? telemetryObserver = null) =>
        new(
            factory,
            new StringReader(input),
            output,
            error,
            isOutputRedirected: true,
            daemonConnection,
            latestVersionProvider,
            telemetryObserver);

    private sealed class RecordingDaemonConnection(
        ServiceResponse response,
        DaemonConnectionPolicy.DaemonFailureState failureState)
        : ICliDaemonConnection
    {
        internal List<ServiceRequest> Requests { get; } = [];

        public DaemonConnectionPolicy.DaemonObservation Observe(string pipeName) =>
            new(DaemonRunning: failureState.Running, StartupInProgress: false);

        public Task<ServiceResponse> SendControlRequestAsync(
            string pipeName,
            ServiceRequest request,
            CancellationToken cancellationToken,
            TimeSpan timeout)
        {
            Requests.Add(request);
            return Task.FromResult(response);
        }

        public DaemonConnectionPolicy.DaemonFailureState ResolveFailureState(
            string pipeName,
            ServiceResponse serviceResponse) =>
            failureState;
    }

    private sealed class RecordingClientFactory(params ServiceResponse[] responses)
        : ICliRequestClientFactory
    {
        private readonly Queue<ServiceResponse> _responses = new(responses);

        internal List<ServiceRequest> Requests { get; } = [];
        internal int ConnectionCount { get; private set; }
        internal Task<ServiceResponse>? PendingResponse { get; init; }
        internal Func<ServiceRequest, Task<ServiceResponse>>? Responder { get; init; }
        internal TaskCompletionSource RequestSent { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken)
        {
            ConnectionCount++;
            return Task.FromResult<ICliRequestClient>(new RecordingClient(this));
        }

        private sealed class RecordingClient(RecordingClientFactory owner) : ICliRequestClient
        {
            public Task<ServiceResponse> SendAsync(
                ServiceRequest request,
                CancellationToken cancellationToken)
            {
                owner.Requests.Add(request);
                owner.RequestSent.TrySetResult();
                return owner.Responder?.Invoke(request)
                    ?? owner.PendingResponse ?? Task.FromResult(owner._responses.Dequeue());
            }

            public void Dispose()
            {
            }
        }
    }
}
