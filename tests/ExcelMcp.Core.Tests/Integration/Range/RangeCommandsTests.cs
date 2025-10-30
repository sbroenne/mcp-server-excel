using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;

/// <summary>
/// Integration tests for RangeCommands - main partial class with shared fixture
/// Other test methods are in partial files: Values.cs, Formulas.cs, Editing.cs, Search.cs, Discovery.cs, Hyperlinks.cs
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public partial class RangeCommandsTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly RangeCommands _commands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles = new();

    public RangeCommandsTests(ITestOutputHelper output)
    {
        _output = output;
        _commands = new RangeCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelMcpRangeTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        foreach (var file in _createdFiles)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch
            {
                // Cleanup is best-effort
            }
        }

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup is best-effort
        }

        GC.SuppressFinalize(this);
    }

    private string CreateTestWorkbook(string name = "test.xlsx")
    {
        string path = Path.Combine(_tempDir, name);
        var fileCommands = new FileCommands();
        var task = Task.Run(async () =>
        {
            await fileCommands.CreateEmptyAsync(path);
        });
        task.Wait();
        _createdFiles.Add(path);
        return path;
    }
}
