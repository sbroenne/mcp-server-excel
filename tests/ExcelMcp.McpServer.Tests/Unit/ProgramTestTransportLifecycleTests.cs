using System.IO.Pipelines;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ProgramTransport")]
[Trait("Speed", "Fast")]
public sealed class ProgramTestTransportLifecycleTests : IDisposable
{
    public void Dispose()
    {
        Program.ResetTestTransport();
    }

    [Fact]
    public void RequestTestTransportShutdown_CanBeCalledMultipleTimesBeforeReset()
    {
        Program.ConfigureTestTransport(new Pipe(), new Pipe());

        Program.RequestTestTransportShutdown();
        var exception = Record.Exception(Program.RequestTestTransportShutdown);

        Assert.Null(exception);
    }

    [Fact]
    public void ResetTestTransport_AfterShutdown_AllowsFreshConfigure()
    {
        Program.ConfigureTestTransport(new Pipe(), new Pipe());

        Program.RequestTestTransportShutdown();
        Program.ResetTestTransport();

        var exception = Record.Exception(() => Program.ConfigureTestTransport(new Pipe(), new Pipe()));

        try
        {
            Assert.Null(exception);
        }
        finally
        {
            Program.ResetTestTransport();
        }
    }
}
