using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Safety;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "SafetyDurability")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DurableFileWriterTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-durable-writer-{Guid.NewGuid():N}");

    [Fact]
    public void WriteUtf8Atomically_CreatesThenReplacesParseableRecordWithoutTempArtifacts()
    {
        Directory.CreateDirectory(_root);
        var path = Path.Combine(_root, "operation.json");

        DurableFileWriter.WriteUtf8Atomically(path, "{\"version\":1}");
        DurableFileWriter.WriteUtf8Atomically(path, "{\"version\":2,\"complete\":true}");

        using var document = JsonDocument.Parse(File.ReadAllText(path));
        Assert.Equal(2, document.RootElement.GetProperty("version").GetInt32());
        Assert.True(document.RootElement.GetProperty("complete").GetBoolean());
        Assert.Empty(Directory.EnumerateFiles(_root, "*.tmp", SearchOption.TopDirectoryOnly));
    }

    [Fact]
    public void WriteUtf8Atomically_WhenPublishFails_PreservesPriorRecordAndCleansTemp()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        Directory.CreateDirectory(_root);
        var path = Path.Combine(_root, "operation.json");
        File.WriteAllText(path, "{\"version\":1}");
        using var destinationLock = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        Assert.ThrowsAny<IOException>(() =>
            DurableFileWriter.WriteUtf8Atomically(path, "{\"version\":2}"));

        Assert.Equal("{\"version\":1}", File.ReadAllText(path));
        Assert.Empty(Directory.EnumerateFiles(_root, "*.tmp", SearchOption.TopDirectoryOnly));
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsm")]
    [InlineData(".xls")]
    public void PendingCheckpointPath_StaysBesideDestinationAndPreservesExcelExtension(string extension)
    {
        var destination = Path.Combine(_root, $"recovery{extension}");

        var pending = WorkbookCheckpointManager.GetPendingCheckpointPath(destination);

        Assert.Equal(Path.GetDirectoryName(destination), Path.GetDirectoryName(pending));
        Assert.Equal(extension, Path.GetExtension(pending));
        Assert.NotEqual(destination, pending);
        Assert.Contains("pending", Path.GetFileName(pending), StringComparison.OrdinalIgnoreCase);
    }

    public void Dispose()
    {
        if (Directory.Exists(_root))
        {
            Directory.Delete(_root, recursive: true);
        }
    }
}
