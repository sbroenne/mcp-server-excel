using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for File Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
///
/// WHAT LLMs NEED TO KNOW:
/// 1. TestFile returns metadata (Exists, IsValid, Message) without Success flag
/// 2. File creation uses SessionManager.CreateSessionForNewFile (create action)
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test Excel COM file operations and Result objects
/// - ✅ Test business rules (valid extensions, file metadata)
/// - ❌ DO NOT test CLI argument parsing (CLI's responsibility)
/// - ❌ DO NOT test JSON serialization (MCP Server's responsibility)
/// - ❌ DO NOT test infrastructure (paths, directories, OS validation)
/// </summary>
[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Files")]
[Trait("RequiresExcel", "false")]
public sealed partial class ServiceFileCommandsTests :
    IClassFixture<ServiceFileTestFixture>
{
    private readonly ServiceFileCommands _fileCommands;
    private readonly ServiceFileTestFixture _fixture;

    public ServiceFileCommandsTests(ServiceFileTestFixture fixture)
    {
        _fileCommands = fixture.Commands;
        _fixture = fixture;
    }
}

public sealed class ServiceFileTestFixture : IDisposable
{
    private readonly string _tempDir =
        Path.Combine(Path.GetTempPath(), $"ServiceFileTests_{Guid.NewGuid():N}");
    private readonly ExcelMcpService _service = new();

    public ServiceFileTestFixture()
    {
        Directory.CreateDirectory(_tempDir);
        Commands = new ServiceFileCommands(_service);
    }

    internal ServiceFileCommands Commands { get; }
    internal string TempDir => _tempDir;

    internal string CreateTestFile()
    {
        var source = Path.Combine(
            AppContext.BaseDirectory,
            "TestFiles",
            "batch-test-static.xlsx");
        var destination = Path.Combine(_tempDir, $"{Guid.NewGuid():N}.xlsx");
        File.Copy(source, destination);
        return destination;
    }

    public void Dispose()
    {
        _service.Dispose();
        if (Directory.Exists(_tempDir))
        {
            Directory.Delete(_tempDir, recursive: true);
        }
    }
}

internal sealed class ServiceFileCommands(ExcelMcpService service)
{
    internal FileValidationInfo Test(string filePath)
    {
        var response = service.ProcessAsync(new ServiceRequest
        {
            Command = "session.test",
            Args = JsonSerializer.Serialize(
                new { filePath },
                ServiceProtocol.JsonOptions),
            Source = "service-file-tests"
        }).GetAwaiter().GetResult();

        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        Assert.False(string.IsNullOrWhiteSpace(response.Result));
        return JsonSerializer.Deserialize<FileValidationInfo>(
                response.Result,
                ServiceProtocol.JsonOptions)
            ?? throw new InvalidOperationException(
                "session.test returned no file validation result.");
    }
}
