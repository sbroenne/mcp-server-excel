namespace Sbroenne.ExcelMcp.Telemetry;

internal sealed record UsageEvent(string Name)
{
    internal Dictionary<string, string> Properties { get; } = [];
}

internal sealed record UsageOperation
{
    internal required string Name { get; init; }
    internal DateTimeOffset Timestamp { get; init; }
    internal TimeSpan Duration { get; init; }
    internal bool Success { get; init; }
    internal Dictionary<string, string> Properties { get; } = [];
}
