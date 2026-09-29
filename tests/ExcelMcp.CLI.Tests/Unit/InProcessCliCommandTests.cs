using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
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
        Assert.True(root.GetProperty("isError").GetBoolean());
        Assert.Equal(
            "Workbook is already open; close the file and retry with exclusive access.",
            root.GetProperty("error").GetString());
        Assert.Equal(
            "Workbook is already open; close the file and retry with exclusive access.",
            root.GetProperty("errorMessage").GetString());
        Assert.Equal("ComInterop", root.GetProperty("errorCategory").GetString());
        Assert.Equal(nameof(IOException), root.GetProperty("exceptionType").GetString());
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
    public async Task SessionCommand_MapsShowFlagAndDefaults(
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
            show
                ? """{"filePath":"C:\\workbooks\\book.xlsx","show":true}"""
                : """{"filePath":"C:\\workbooks\\book.xlsx","show":false}""",
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
        Assert.Contains(
            "Excel automation powered by ExcelMcp Core",
            error.ToString(),
            StringComparison.Ordinal);
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
        Assert.Equal(expectedState, envelope.RootElement.GetProperty("daemonState").GetString());
        if (running)
        {
            Assert.True(envelope.RootElement.GetProperty("running").GetBoolean());
            Assert.Equal(
                "controlled transport failure",
                envelope.RootElement.GetProperty("error").GetString());
        }
        else
        {
            Assert.Equal(0, envelope.RootElement.GetProperty("count").GetInt32());
            Assert.Empty(envelope.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.False(envelope.RootElement.TryGetProperty("running", out _));
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

        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken)
        {
            return Task.FromResult<ICliRequestClient>(new RecordingClient(this));
        }

        private sealed class RecordingClient(RecordingClientFactory owner) : ICliRequestClient
        {
            public Task<ServiceResponse> SendAsync(
                ServiceRequest request,
                CancellationToken cancellationToken)
            {
                owner.Requests.Add(request);
                return Task.FromResult(owner._responses.Dequeue());
            }

            public void Dispose()
            {
            }
        }
    }
}
