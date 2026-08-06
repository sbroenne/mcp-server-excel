using System.Threading.Channels;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "ExcelBatch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelBatchQueuePolicyTests
{
    [Fact]
    public void WorkbookQueue_IsFiniteAndUsesWaitBackpressureWithoutDrops()
    {
        Assert.InRange(ExcelBatch.WorkQueueCapacity, 1, 64);
        Assert.Equal(BoundedChannelFullMode.Wait, ExcelBatch.WorkQueueFullMode);
    }
}
