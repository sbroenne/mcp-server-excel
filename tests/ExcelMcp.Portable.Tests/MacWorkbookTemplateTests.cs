using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacWorkbookTemplateTests
{
    [Fact]
    public void CopyCreatesExactOpaqueWorkbook()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "created.xlsx");
        try
        {
            MacWorkbookTemplate.Copy(path, macroEnabled: false);

            Assert.True(File.Exists(path));
            using var expected = typeof(MacWorkbookTemplate).Assembly
                .GetManifestResourceStream(
                    "Sbroenne.ExcelMcp.Service.Mac.Blank.xlsx");
            Assert.NotNull(expected);
            using var actual = File.OpenRead(path);
            Assert.True(expected!.Length > 0);
            Assert.Equal(expected.Length, actual.Length);
            Assert.True(StreamContentsEqual(expected, actual));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void CopyNeverOverwritesExistingDestination()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "existing.xlsx");
        try
        {
            File.WriteAllText(path, "sentinel");

            Assert.Throws<IOException>(() =>
                MacWorkbookTemplate.Copy(path, macroEnabled: false));

            Assert.Equal("sentinel", File.ReadAllText(path));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp-template-*.tmp"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void CopyRejectsMacroEnabledCreationWithoutCreatingFile()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "created.xlsm");
        try
        {
            var error = Assert.Throws<PlatformNotSupportedException>(() =>
                MacWorkbookTemplate.Copy(path, macroEnabled: true));

            Assert.Contains("Excel-authored .xlsm template", error.Message);
            Assert.False(File.Exists(path));
            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string CreateTemporaryDirectory()
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"ExcelMcp-MacWorkbookTemplate-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static bool StreamContentsEqual(Stream expected, Stream actual)
    {
        var expectedBuffer = new byte[81920];
        var actualBuffer = new byte[81920];
        while (true)
        {
            int expectedRead = expected.Read(expectedBuffer);
            int actualRead = actual.Read(actualBuffer);
            if (expectedRead != actualRead)
            {
                return false;
            }
            if (expectedRead == 0)
            {
                return true;
            }
            if (!expectedBuffer.AsSpan(0, expectedRead)
                    .SequenceEqual(actualBuffer.AsSpan(0, actualRead)))
            {
                return false;
            }
        }
    }
}
