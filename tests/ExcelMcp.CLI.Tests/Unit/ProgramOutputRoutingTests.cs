using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ProgramOutput")]
[Trait("Speed", "Fast")]
public sealed class ProgramOutputRoutingTests
{
    [Fact]
    public void WriteDiagnosticMarkupLine_WritesToStandardErrorOnly()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var originalOut = Console.Out;
        var originalError = Console.Error;

        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);

            Program.WriteDiagnosticMarkupLine("[red]diagnostic[/]");
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }

        Assert.Empty(stdout.ToString());
        Assert.Contains("diagnostic", stderr.ToString(), StringComparison.Ordinal);
    }
}
