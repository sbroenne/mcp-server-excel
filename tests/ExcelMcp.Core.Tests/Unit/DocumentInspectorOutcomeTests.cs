using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "DocumentInspector")]
[Trait("RequiresExcel", "false")]
public sealed class DocumentInspectorOutcomeTests
{
    [Theory]
    [InlineData(0, "ok")]
    [InlineData(1, "issue-found")]
    [InlineData(2, "error")]
    [InlineData(-1, "unknown")]
    [InlineData(99, "unknown")]
    public void NativeStatus_PreservesCodeAndLocalizedText(int code, string label)
    {
        var outcome = new DocumentInspectorOutcome(code, "Native result");
        var finding = new DocumentInspectorFinding(3, "Native &name", code, "Native result");
        Assert.Equal(code, outcome.StatusCode);
        Assert.Equal(label, outcome.Status);
        Assert.Equal(label, finding.Status);
        Assert.Equal("Native result", outcome.Results);
        Assert.Equal("Native &name", finding.Name);
    }

    [Fact]
    public void Inspection_NoModulesIsAnExplicitFailure()
    {
        var result = WorkbookCommands.CreateDocumentInspectionResult([]);
        Assert.False(result.Success);
        Assert.Empty(result.Inspectors);
        Assert.Contains("No document inspection was performed", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("Comments/revisions and document properties are not included", result.ReadLimitations,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Inspection_CompletedChecksSucceedEvenWithFindings(int code)
    {
        var finding = new DocumentInspectorFinding(1, "Module", code, "Native result");
        var result = WorkbookCommands.CreateDocumentInspectionResult([finding]);
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(finding, Assert.Single(result.Inspectors));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(99)]
    public void Inspection_NativeFailureRetainsEveryModuleAndErrorContext(int code)
    {
        var completed = new DocumentInspectorFinding(1, "Completed", 1, "Findings retained");
        var failed = new DocumentInspectorFinding(2, "Failed", code, "Native failure");
        var later = new DocumentInspectorFinding(3, "Later", 0, "No findings");
        var result = WorkbookCommands.CreateDocumentInspectionResult([completed, failed, later]);
        Assert.False(result.Success);
        Assert.Equal([completed, failed, later], result.Inspectors);
        Assert.Contains($"'Failed' (index 2, status {code}): Native failure", result.ErrorMessage,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 0, true)]
    [InlineData(1, 0, false)]
    [InlineData(2, 0, false)]
    [InlineData(0, 1, false)]
    [InlineData(0, 2, false)]
    [InlineData(99, 0, false)]
    [InlineData(0, 99, false)]
    public void Fix_RequiresBothSuccessfulFixAndCleanReinspection(int fixCode, int inspectionCode, bool success)
    {
        var fix = new DocumentInspectorOutcome(fixCode, "Native fix result");
        var inspection = new DocumentInspectorOutcome(inspectionCode, "Native inspection result");
        var result = WorkbookCommands.CreateDocumentInspectorFixResult(3, "Selected", fix, inspection);
        Assert.Equal(success, result.Success);
        Assert.Equal(3, result.InspectorIndex);
        Assert.Equal("Selected", result.InspectorName);
        Assert.Equal(fix, result.Fix);
        Assert.Equal(inspection, result.Inspection);
        if (success)
        {
            Assert.Null(result.ErrorMessage);
            Assert.Contains("Saving remains explicit", result.Message, StringComparison.Ordinal);
        }
        else
        {
            Assert.Null(result.Message);
            Assert.Contains($"Fix status {fixCode}: Native fix result", result.ErrorMessage, StringComparison.Ordinal);
            Assert.Contains($"reinspection status {inspectionCode}: Native inspection result", result.ErrorMessage,
                StringComparison.Ordinal);
            Assert.Contains("may already have changed", result.ErrorMessage, StringComparison.Ordinal);
        }
    }
}
