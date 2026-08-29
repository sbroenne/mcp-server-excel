using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "Batch")]
[Trait("Speed", "Fast")]
public sealed class BatchExecutionCancellationTests
{
    [Fact]
    public void Push_NestedScopes_RestorePreviousToken()
    {
        using var outerSource = new CancellationTokenSource();
        using var innerSource = new CancellationTokenSource();

        Assert.False(BatchExecutionCancellation.Current.CanBeCanceled);
        using (BatchExecutionCancellation.Push(
            outerSource.Token,
            requiresCooperativeCleanup: true))
        {
            Assert.Equal(outerSource.Token, BatchExecutionCancellation.Current);
            Assert.True(BatchExecutionCancellation.RequiresCooperativeCleanup);
            using (BatchExecutionCancellation.Push(innerSource.Token))
            {
                Assert.Equal(innerSource.Token, BatchExecutionCancellation.Current);
                Assert.False(BatchExecutionCancellation.RequiresCooperativeCleanup);
            }

            Assert.Equal(outerSource.Token, BatchExecutionCancellation.Current);
            Assert.True(BatchExecutionCancellation.RequiresCooperativeCleanup);
        }

        Assert.False(BatchExecutionCancellation.Current.CanBeCanceled);
        Assert.False(BatchExecutionCancellation.RequiresCooperativeCleanup);
    }
}
