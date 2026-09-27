using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class PlatformHostTests
{
    [Fact]
    public void McpPathValidation_ReportsTheCurrentPlatform()
    {
        var result = Sbroenne.ExcelMcp.McpServer.Tools.ExcelToolsBase.ValidateWindowsPath("relative.xlsx");
        Assert.NotNull(result);
        using var document = JsonDocument.Parse(result);
        var error = document.RootElement.GetProperty("errorMessage").GetString();
        if (OperatingSystem.IsWindows())
        {
            Assert.Contains("not an absolute Windows path", error, StringComparison.Ordinal);
        }
        else
        {
            Assert.Contains("not an absolute path", error, StringComparison.Ordinal);
        }
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "PlatformHost")]
    public void NormalizeAbsolutePath_RejectsRelativePathWithPlatformNeutralMessage()
    {
        var exception = Assert.Throws<ArgumentException>(
            () => FilePathValidation.NormalizeAbsolutePath("relative/workbook.xlsx"));

        Assert.Contains("absolute path", exception.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Windows", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "PlatformHost")]
    public void NormalizeAbsolutePath_NormalizesCurrentPlatformPath()
    {
        var expected = Path.GetFullPath(
            Path.Combine(Path.GetTempPath(), "excelmcp", "..", "workbook.xlsx"));

        var actual = FilePathValidation.NormalizeAbsolutePath(
            Path.Combine(Path.GetTempPath(), "excelmcp", "..", "workbook.xlsx"));

        Assert.Equal(expected, actual);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "PlatformHost")]
    public async Task ServiceStatus_UsesPortableHostWithoutStartingExcel()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "service.status"
        });

        Assert.True(response.Success);
        Assert.Null(response.ErrorMessage);
        using var result = JsonDocument.Parse(Assert.IsType<string>(response.Result));
        Assert.True(result.RootElement.GetProperty("running").GetBoolean());
        Assert.Equal(0, result.RootElement.GetProperty("sessionCount").GetInt32());
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "PlatformHost")]
    public void PipeNames_AreStableAndUserScoped()
    {
        var first = ServiceSecurity.GetCliPipeName();
        var second = ServiceSecurity.GetCliPipeName();

        Assert.Equal(first, second);
        Assert.StartsWith("excelmcp-cli-", first, StringComparison.Ordinal);
        Assert.DoesNotContain(Path.DirectorySeparatorChar, first);
    }
}
