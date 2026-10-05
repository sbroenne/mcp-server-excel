using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedTestReportTests
{
    [Theory]
    [InlineData(1, 1, 1, "Completed", "Passed", true)]
    [InlineData(0, 0, 0, "Completed", "Passed", false)]
    [InlineData(1, 0, 1, "Completed", "Failed", false)]
    [InlineData(1, 0, 0, "Completed", "NotExecuted", false)]
    [InlineData(1, 1, 1, "Failed", "Passed", false)]
    [InlineData(2, 2, 2, "Completed", "Passed", false)]
    public void Reports_RequireEverySelectedCaseToPass(int total, int passed, int executed, string summary, string outcome, bool succeeds)
    {
        var file = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Report.{Guid.NewGuid():N}.trx");
        try
        {
            File.WriteAllText(file, $$"""
                <TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
                  <Results><UnitTestResult outcome="{{outcome}}"/></Results>
                  <ResultSummary outcome="{{summary}}">
                    <Counters total="{{total}}" passed="{{passed}}" executed="{{executed}}"/>
                  </ResultSummary>
                </TestRun>
                """);
            if (succeeds) { Assert.Equal(total, TestReport.Verify(file)); }
            else { Assert.Throws<InvalidOperationException>(() => TestReport.Verify(file)); }
        }
        finally { File.Delete(file); }
    }
}
