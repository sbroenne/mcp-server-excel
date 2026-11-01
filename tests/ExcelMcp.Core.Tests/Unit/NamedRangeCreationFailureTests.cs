using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "false")]
public sealed class NamedRangeCreationFailureTests
{
    [Theory]
    [InlineData("=0")]
    [InlineData("=SUM(1.25,2.5)")]
    public void ReportsActualReferenceAndPreservesOriginalFailure(string reference)
    {
        var cause = new InvalidOperationException("Assigning the reference failed.");
        var failure = NamedRangeCommands.DescribeCreatedNameFailure("InputRate", cause, () => reference);

        Assert.Same(cause, failure.InnerException);
        Assert.Contains("InputRate", failure.Message);
        Assert.Contains(reference, failure.Message);
        Assert.Contains("not deleted", failure.Message);
        Assert.Contains(cause.Message, failure.Message);
    }

    [Fact]
    public void InspectionFailureIsRetainedAlongsideOriginalFailure()
    {
        var cause = new InvalidOperationException("Assigning the reference failed.");
        var inspection = new InvalidOperationException("Reading the reference failed.");
        var failure = NamedRangeCommands.DescribeCreatedNameFailure("InputRate", cause, () => throw inspection);

        var aggregate = Assert.IsType<AggregateException>(failure.InnerException);
        Assert.Equal([cause, inspection], aggregate.InnerExceptions);
        Assert.Contains("could not be inspected", failure.Message);
        Assert.Contains("not deleted", failure.Message);
        Assert.Contains(cause.Message, failure.Message);
        Assert.Contains(inspection.Message, failure.Message);
    }
}
