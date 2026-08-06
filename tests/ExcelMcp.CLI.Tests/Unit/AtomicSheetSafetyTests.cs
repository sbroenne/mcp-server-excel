using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class AtomicSheetSafetyTests : IDisposable
{
    private readonly string _stateRoot = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-atomic-sheet-safety-{Guid.NewGuid():N}");

    [Theory]
    [InlineData("sheet.copy-to-file", true, null, false)]
    [InlineData("sheet.copy-to-file", false, "review-123", false)]
    [InlineData("sheet.move-to-file", false, null, true)]
    public async Task AtomicCrossFileMutation_WithSafetyOption_FailsClosedBeforeExcelDispatch(
        string command,
        bool reviewOnly,
        string? reviewId,
        bool checkpoint)
    {
        using var service = new ExcelMcpService(_stateRoot);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            Args = "{}",
            ReviewOnly = reviewOnly,
            ReviewId = reviewId,
            Checkpoint = checkpoint
        });

        Assert.False(response.Success);
        Assert.Equal("SafetyWorkflowUnavailable", response.ErrorCategory);
        Assert.Contains("neither workbook was changed", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    public void Dispose()
    {
        if (Directory.Exists(_stateRoot))
        {
            Directory.Delete(_stateRoot, recursive: true);
        }

        GC.SuppressFinalize(this);
    }
}
