using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceClient")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceClientDeadlineTests
{
    [Fact]
    public void GetStepTimeout_UsesOneControlledTotalBudgetAcrossSteps()
    {
        var clock = new AdvancingTimeProvider();
        var startedAt = clock.GetTimestamp();

        clock.Advance(TimeSpan.FromSeconds(3));
        Assert.Equal(
            TimeSpan.FromSeconds(5),
            ServiceClient.GetStepTimeout(
                TimeSpan.FromSeconds(5),
                TimeSpan.FromSeconds(10),
                startedAt,
                clock));

        clock.Advance(TimeSpan.FromSeconds(4));
        Assert.Equal(
            TimeSpan.FromSeconds(3),
            ServiceClient.GetStepTimeout(
                TimeSpan.FromSeconds(5),
                TimeSpan.FromSeconds(10),
                startedAt,
                clock));

        clock.Advance(TimeSpan.FromSeconds(3));
        Assert.Throws<TimeoutException>(() =>
            ServiceClient.GetStepTimeout(
                TimeSpan.FromSeconds(5),
                TimeSpan.FromSeconds(10),
                startedAt,
                clock));
    }

    private sealed class AdvancingTimeProvider : TimeProvider
    {
        private long _timestamp;

        public override long TimestampFrequency => TimeSpan.TicksPerSecond;

        public override long GetTimestamp() => _timestamp;

        internal void Advance(TimeSpan elapsed) => _timestamp += elapsed.Ticks;
    }
}
