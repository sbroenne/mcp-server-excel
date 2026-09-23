using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Tests for FileCommands TestFile operation
/// </summary>
public partial class FileCommandsTests
{
    [Fact]
    public void Test_ExistingValidFile_ReturnsSuccess()
    {
        // Arrange - Create a valid file
        var testFile = _fixture.CreateTestFile();

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.True(info.Success);
        Assert.True(info.CanOpen);
        Assert.False(info.WillOpenReadOnly);
        Assert.False(info.RequiresVisibleSession);
        Assert.Equal(".xlsx", info.Extension);
        Assert.True(info.Size > 0);
        Assert.Null(info.Message);
    }
    [Fact]
    public void Test_NonExistent_ReturnsFailure()
    {
        // Arrange
        string testFile = Path.Join(_fixture.TempDir, $"NonExistent_{Guid.NewGuid():N}.xlsx");

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.False(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.Null(info.IsError);
        Assert.NotNull(info.Message);
        Assert.Contains("not found", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_CorruptSupportedExtension_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"Corrupt_{Guid.NewGuid():N}.xlsx");
        System.IO.File.WriteAllText(testFile, "not an Excel workbook");

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("valid Excel workbook", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_ZipWithWorkbookEntryNamesButInvalidXml_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"CorruptParts_{Guid.NewGuid():N}.xlsx");
        using (var archive = System.IO.Compression.ZipFile.Open(
                   testFile,
                   System.IO.Compression.ZipArchiveMode.Create))
        {
            foreach (var name in new[] { "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml" })
            {
                var entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write("<invalid");
            }
        }

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.CanOpen);
    }

    [Fact]
    public void Test_ZipWithFabricatedWorkbookNamespaces_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"FabricatedParts_{Guid.NewGuid():N}.xlsx");
        using (var archive = System.IO.Compression.ZipFile.Open(
                   testFile,
                   System.IO.Compression.ZipArchiveMode.Create))
        {
            WriteZipEntry(archive, "[Content_Types].xml", "<Types/>");
            WriteZipEntry(archive, "_rels/.rels",
                """<Relationships><Relationship Id="rId1" Type="x/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
            WriteZipEntry(archive, "xl/workbook.xml",
                """<workbook><sheets><sheet name="Sheet1" sheetId="1" id="rId1"/></sheets></workbook>""");
            WriteZipEntry(archive, "xl/_rels/workbook.xml.rels",
                """<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");
            WriteZipEntry(archive, "xl/worksheets/sheet1.xml", "<worksheet/>");
        }

        var info = _fileCommands.Test(testFile);

        Assert.False(info.IsValid);
        Assert.False(info.CanOpen);
    }

    [Fact]
    public void Test_OoxmlPackageWithInvalidWorksheetPart_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"InvalidWorksheet_{Guid.NewGuid():N}.xlsx");
        using (var archive = System.IO.Compression.ZipFile.Open(
                   testFile,
                   System.IO.Compression.ZipArchiveMode.Create))
        {
            WriteZipEntry(archive, "[Content_Types].xml",
                """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");
            WriteZipEntry(archive, "_rels/.rels",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
            WriteZipEntry(archive, "xl/workbook.xml",
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""");
            WriteZipEntry(archive, "xl/_rels/workbook.xml.rels",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");
            WriteZipEntry(archive, "xl/worksheets/sheet1.xml",
                """<notWorksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""");
        }

        var info = _fileCommands.Test(testFile);

        Assert.False(info.IsValid);
        Assert.False(info.CanOpen);
    }

    [Fact]
    public void Test_OoxmlPackageWithInvalidSecondSheet_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"InvalidSecondSheet_{Guid.NewGuid():N}.xlsx");
        using (var archive = System.IO.Compression.ZipFile.Open(
                   testFile,
                   System.IO.Compression.ZipArchiveMode.Create))
        {
            WriteZipEntry(archive, "[Content_Types].xml",
                """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");
            WriteZipEntry(archive, "_rels/.rels",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
            WriteZipEntry(archive, "xl/workbook.xml",
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>""");
            WriteZipEntry(archive, "xl/_rels/workbook.xml.rels",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" TargetMode="External" Target="https://example.invalid/sheet2.xml"/></Relationships>""");
            WriteZipEntry(archive, "xl/worksheets/sheet1.xml",
                """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""");
        }

        var info = _fileCommands.Test(testFile);

        Assert.False(info.IsValid);
        Assert.False(info.CanOpen);
    }

    private static void WriteZipEntry(
        System.IO.Compression.ZipArchive archive,
        string name,
        string content)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    [Fact]
    public void Test_LockedSupportedFile_ReportsNotOpenable()
    {
        var testFile = _fixture.CreateTestFile();
        using var lockStream = new FileStream(
            testFile,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None);

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("already open", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("TestFile.xls", ".xls")]
    [InlineData("TestFile.csv", ".csv")]
    [InlineData("TestFile.txt", ".txt")]
    public void Test_InvalidExtension_ReturnsFailure(string fileName, string expectedExt)
    {
        // Arrange
        string testFile = Path.Join(_fixture.TempDir, $"{Guid.NewGuid():N}_{fileName}");

        // Create file with invalid extension
        System.IO.File.WriteAllText(testFile, "test content");

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.Equal(expectedExt, info.Extension);
        Assert.NotNull(info.Message);
        Assert.Contains("Invalid file extension", info.Message, StringComparison.OrdinalIgnoreCase);
    }
}
