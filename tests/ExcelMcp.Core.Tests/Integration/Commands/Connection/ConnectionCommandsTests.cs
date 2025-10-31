using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Comprehensive integration tests for ConnectionCommands.
/// Tests all connection operations with batch API pattern.
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests : IDisposable
{
    private readonly ConnectionCommands _commands;
    private readonly string _tempDir;
    private bool _disposed;

    public ConnectionCommandsTests()
    {
        _commands = new ConnectionCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_Conn_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Helper to create a test CSV file for text connections
    /// </summary>
    private string CreateTestCsvFile(string fileName = "data.csv")
    {
        var csvFile = Path.Combine(_tempDir, fileName);
        File.WriteAllText(csvFile, "Name,Value\nTest1,100\nTest2,200");
        return csvFile;
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch { /* Cleanup failure is non-critical */ }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
