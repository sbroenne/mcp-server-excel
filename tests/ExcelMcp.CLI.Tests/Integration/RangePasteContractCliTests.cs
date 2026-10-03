using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangePasteContractCliTests
{
    public static TheoryData<string, bool, bool> PasteOptions
    {
        get
        {
            TheoryData<string, bool, bool> data = [];
            foreach (var kind in new[] { "all", "values", "formulas", "formats", "validation" })
                foreach (var transpose in new[] { false, true })
                    foreach (var skipBlanks in new[] { false, true })
                        data.Add(kind, transpose, skipBlanks);
            return data;
        }
    }

    [Theory]
    [MemberData(nameof(PasteOptions))]
    public async Task Copy_MapsNativePasteOptionsAndReturnsResolvedBounds(string pasteKind, bool transpose, bool skipBlanks)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "copy", "--session", "session-1",
            "--source-sheet", "Source", "--source-range", "A1:B3",
            "--target-sheet", "Target", "--target-range", "D1",
            "--paste-kind", pasteKind, "--transpose", transpose ? "true" : "false",
            "--skip-blanks", skipBlanks ? "true" : "false"
        ], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    destinationAddress = transpose ? "$D$1:$F$2" : "$D$1:$E$3",
                    pasteKind,
                    transpose,
                    skipBlanks
                })
            };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("range.copy", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(pasteKind, args.RootElement.GetProperty("pasteKind").GetString(), ignoreCase: true);
        Assert.Equal(transpose, args.RootElement.GetProperty("transpose").GetBoolean());
        Assert.Equal(skipBlanks, args.RootElement.GetProperty("skipBlanks").GetBoolean());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal(transpose ? "$D$1:$F$2" : "$D$1:$E$3",
            output.RootElement.GetProperty("destinationAddress").GetString());
    }

    [Theory]
    [InlineData("copy-values")]
    [InlineData("copy-formulas")]
    public async Task RemovedCopyAction_IsRejected(string action)
    {
        var result = await InProcessCliHelper.RunAsync(["range", action, "--session", "session-1"]);
        Assert.NotEqual(0, result.ExitCode);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("unknown")]
    [InlineData("99")]
    public async Task Copy_MissingOrInvalidKindDoesNotDispatch(string? pasteKind)
    {
        List<string> arguments =
        [
            "range", "copy", "--session", "session-1",
            "--source-sheet", "Source", "--source-range", "A1:B3",
            "--target-sheet", "Target", "--target-range", "D1"
        ];
        if (pasteKind is not null)
            arguments.AddRange(["--paste-kind", pasteKind]);
        var result = await InProcessCliHelper.RunAsync(arguments);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("pasteKind", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }
}
