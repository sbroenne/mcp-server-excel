namespace Sbroenne.ExcelMcp.ComInterop.Tests.Helpers;

public sealed class TempDirectoryFixture : IDisposable
{
    public TempDirectoryFixture()
    {
        DirectoryPath = Path.Combine(
            Path.GetTempPath(),
            $"ExcelMcpFileValidation_{Guid.NewGuid():N}");
        Directory.CreateDirectory(DirectoryPath);
    }

    public string DirectoryPath { get; }

    public string CreateFilePath() =>
        Path.Combine(DirectoryPath, $"{Guid.NewGuid():N}.xlsx");

    public void Dispose()
    {
        if (Directory.Exists(DirectoryPath))
        {
            Directory.Delete(DirectoryPath, recursive: true);
        }
    }
}
