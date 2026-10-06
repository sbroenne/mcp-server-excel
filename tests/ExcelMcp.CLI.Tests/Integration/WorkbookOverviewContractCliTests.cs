using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Workbook")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookOverviewContractCliTests
{
    [Fact]
    public async Task Inspect_MapsSelectionSectionsAndLimits()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "-q", "workbook", "inspect", "--session", "session-1",
            "--sheet-name", "Summary",
            "--include-sheets", "true",
            "--include-tables", "false",
            "--include-defined-names", "true",
            "--include-preview", "true",
            "--range-address", "A1:D20",
            "--max-items", "7",
            "--max-preview-rows", "3",
            "--max-preview-columns", "4",
            "--max-cell-characters", "20",
            "--max-preview-characters", "123"
        ], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"preview":{"rangeAddress":"$A$1:$D$3"}}"""
            };
        });

        Assert.Equal(0, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal("workbook.inspect", captured.Command);
        using var arguments = JsonDocument.Parse(captured.Args!);
        var args = arguments.RootElement;
        Assert.Equal("Summary", args.GetProperty("sheetName").GetString());
        Assert.True(args.GetProperty("includeSheets").GetBoolean());
        Assert.False(args.GetProperty("includeTables").GetBoolean());
        Assert.True(args.GetProperty("includeDefinedNames").GetBoolean());
        Assert.True(args.GetProperty("includePreview").GetBoolean());
        Assert.Equal("A1:D20", args.GetProperty("rangeAddress").GetString());
        Assert.Equal(7, args.GetProperty("maxItems").GetInt32());
        Assert.Equal(3, args.GetProperty("maxPreviewRows").GetInt32());
        Assert.Equal(4, args.GetProperty("maxPreviewColumns").GetInt32());
        Assert.Equal(20, args.GetProperty("maxCellCharacters").GetInt32());
        Assert.Equal(123, args.GetProperty("maxPreviewCharacters").GetInt32());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal("$A$1:$D$3", output.RootElement.GetProperty("preview")
            .GetProperty("rangeAddress").GetString());
    }
}
