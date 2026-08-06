// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelFileValidatorTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(
        Path.GetTempPath(),
        $"ExcelFileValidatorTests-{Guid.NewGuid():N}");

    [Fact]
    public void Inspect_LegacyXls_PreservesOpenCompatibilityWithoutChangingFileTestPolicy()
    {
        Directory.CreateDirectory(_tempDirectory);
        var filePath = Path.Combine(_tempDirectory, "legacy.xls");
        File.WriteAllBytes(filePath, [0xD0, 0xCF, 0x11, 0xE0]);

        var result = ExcelFileValidator.Inspect(filePath);

        Assert.True(result.IsOpenableExtension);
        Assert.False(result.IsSupportedExtension);
        Assert.False(result.IsValidExistingWorkbook);
        Assert.Contains("can be opened", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inspect_CreatePathBeyondExcelPracticalLimit_IsRejectedOnlyForCreation()
    {
        Directory.CreateDirectory(_tempDirectory);
        var requiredNameLength = ExcelFileValidator.MaximumCreatePathLength - _tempDirectory.Length + 8;
        var filePath = Path.Combine(_tempDirectory, $"{new string('x', requiredNameLength)}.xlsx");

        var result = ExcelFileValidator.Inspect(filePath);

        Assert.True(result.IsWithinPathLimit);
        Assert.False(result.IsWithinCreatePathLimit);
        Assert.True(result.FilePath.Length > ExcelFileValidator.MaximumCreatePathLength);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDirectory))
        {
            Directory.Delete(_tempDirectory, recursive: true);
        }
    }
}
