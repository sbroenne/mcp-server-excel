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
        using (var args = JsonDocument.Parse(request.Args!))
        {
            Assert.Equal(workbookPath, args.RootElement.GetProperty("filePath").GetString());
            Assert.False(args.RootElement.GetProperty("show").GetBoolean());
            Assert.False(args.RootElement.TryGetProperty("timeoutSeconds", out _));
        }

        using var envelope = JsonDocument.Parse(output.ToString());
        var root = envelope.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("isError").GetBoolean());
        Assert.Equal(root.GetProperty("error").GetString(), root.GetProperty("errorMessage").GetString());
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

    private static CliCommandRuntime CreateRuntime(
        RecordingClientFactory factory,
        StringWriter output,
        StringWriter error,
        string input = "") =>
        new(
            factory,
            new StringReader(input),
            output,
            error,
            isOutputRedirected: true);

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
