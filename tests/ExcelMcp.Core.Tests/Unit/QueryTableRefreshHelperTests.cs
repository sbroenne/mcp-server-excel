using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "QueryTable")]
[Trait("RequiresExcel", "false")]
public sealed class QueryTableRefreshHelperTests
{
    [Fact]
    public void EnsureSucceeded_WhenExcelCancelsRefresh_ThrowsCategorizedFailure()
    {
        var error = Assert.Throws<OperationFailureException>(
            () => QueryTableRefreshHelper.EnsureSucceeded(false, "QueryTable refresh"));

        Assert.Equal(OperationFailureCategory.Cancelled, error.ErrorCategory);
        Assert.Contains("cancelled", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EnsureSucceeded_WhenRefreshCompletes_DoesNotThrow()
    {
        QueryTableRefreshHelper.EnsureSucceeded(true, "QueryTable refresh");
    }
}
