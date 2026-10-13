using Microsoft.ApplicationInsights.DataContracts;
using Sbroenne.ExcelMcp.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.Tests.Telemetry;

[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class TelemetryUserIdentityTests
{
    [Fact]
    public void RoleInstance_UsesOnlyTheRandomUserIdPrefix()
    {
        var identity = new TelemetryUserIdentity("0123456789abcdef0123456789abcdef", null);

        Assert.Equal("instance-01234567", identity.RoleInstance);
    }

    [Fact]
    public void GetOrCreate_CreatesAndReusesOneRandomIdentity()
    {
        using var directory = new TemporaryDirectory();
        var filePath = GetIdentityPath(directory.Path);

        var created = TelemetryUserIdentityStore.GetOrCreate(directory.Path);
        var reused = TelemetryUserIdentityStore.GetOrCreate(directory.Path);

        Assert.True(Guid.TryParseExact(created.UserId, "N", out _));
        Assert.Equal(created.UserId, File.ReadAllText(filePath));
        Assert.Equal(created.UserId, reused.UserId);
        Assert.NotNull(created.LegacyUserId);
        Assert.Null(reused.LegacyUserId);
    }

    [Fact]
    public void GetOrCreate_DeletingTheFileCreatesANewIdentity()
    {
        using var directory = new TemporaryDirectory();
        var filePath = GetIdentityPath(directory.Path);
        var first = TelemetryUserIdentityStore.GetOrCreate(directory.Path);

        File.Delete(filePath);
        var second = TelemetryUserIdentityStore.GetOrCreate(directory.Path);

        Assert.NotEqual(first.UserId, second.UserId);
        Assert.NotNull(second.LegacyUserId);
    }

    [Fact]
    public void GetOrCreate_StorageFailureUsesAProcessOnlyRandomIdentity()
    {
        using var directory = new TemporaryDirectory();
        var unusableAppData = Path.Combine(directory.Path, "app-data");
        Directory.CreateDirectory(directory.Path);
        File.WriteAllText(unusableAppData, "not a directory");

        var identity = TelemetryUserIdentityStore.GetOrCreate(unusableAppData);

        Assert.True(Guid.TryParseExact(identity.UserId, "N", out _));
        Assert.Null(identity.LegacyUserId);
        Assert.False(Directory.Exists(unusableAppData));
    }

    [Fact]
    public void TryDeleteTemporaryFile_StorageFailureDoesNotEscape()
    {
        var exception = Record.Exception(() =>
            TelemetryUserIdentityStore.TryDeleteTemporaryFile(
                "unused",
                _ => throw new NotSupportedException("Storage cleanup is unavailable.")));

        Assert.Null(exception);
    }

    [Fact]
    public async Task GetOrCreate_ConcurrentFirstUseSharesTheCreatedIdentity()
    {
        using var directory = new TemporaryDirectory();

        var identities = await Task.WhenAll(
            Enumerable.Range(0, 8)
                .Select(_ => Task.Run(() => TelemetryUserIdentityStore.GetOrCreate(directory.Path))));

        Assert.Single(identities.Select(identity => identity.UserId).Distinct());
        Assert.Single(identities, identity => identity.LegacyUserId is not null);
    }

    [Fact]
    public void CreateMigrationEvent_UsesNewIdentityAndCarriesLegacyIdentityOnce()
    {
        using var directory = new TemporaryDirectory();
        var identity = TelemetryUserIdentityStore.GetOrCreate(directory.Path);

        EventTelemetry migration = Assert.IsType<EventTelemetry>(
            TelemetryUserIdentityStore.CreateMigrationEvent(identity));

        Assert.Equal("UserIdMigration", migration.Name);
        Assert.Equal(identity.UserId, migration.Context.User.Id);
        Assert.Equal(identity.LegacyUserId, migration.Properties["LegacyUserId"]);
        Assert.Null(TelemetryUserIdentityStore.CreateMigrationEvent(
            TelemetryUserIdentityStore.GetOrCreate(directory.Path)));
    }

    private static string GetIdentityPath(string appDataPath) =>
        Path.Combine(appDataPath, "ExcelMcp", "telemetry-user-id");

    private sealed class TemporaryDirectory : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"excelmcp-telemetry-{Guid.NewGuid():N}");

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
