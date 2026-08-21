using System.IO.Compression;
using Sbroenne.ExcelMcp.ComInterop.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class FileAccessValidatorTests(
    TempDirectoryFixture fixture) : IClassFixture<TempDirectoryFixture>
{
    private const string SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string StrictSpreadsheetNamespace =
        "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string MacroSheetNamespace =
        "http://schemas.microsoft.com/office/excel/2006/main";
    private const string WorksheetRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string StrictWorksheetRelationship =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";
    private const string ChartSheetRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    private const string StrictChartSheetRelationship =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
    private const string MacroSheetRelationship =
        "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
    private const string InternationalMacroSheetRelationship =
        "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
    private const string DialogSheetRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
    private const string StrictDialogSheetRelationship =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
    private const string MacroSheetContentType =
        "application/vnd.ms-excel.macrosheet+xml";
    private const string InternationalMacroSheetContentType =
        "application/vnd.ms-excel.intlmacrosheet+xml";
    private const string DialogSheetContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml";

    [Theory]
    [InlineData(
        WorksheetRelationship,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        "worksheet",
        SpreadsheetNamespace,
        SpreadsheetNamespace)]
    [InlineData(
        StrictWorksheetRelationship,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        "worksheet",
        StrictSpreadsheetNamespace,
        StrictSpreadsheetNamespace)]
    [InlineData(
        ChartSheetRelationship,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
        "chartsheet",
        SpreadsheetNamespace,
        SpreadsheetNamespace)]
    [InlineData(
        StrictChartSheetRelationship,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
        "chartsheet",
        StrictSpreadsheetNamespace,
        StrictSpreadsheetNamespace)]
    [InlineData(
        MacroSheetRelationship,
        MacroSheetContentType,
        "macrosheet",
        MacroSheetNamespace,
        SpreadsheetNamespace)]
    [InlineData(
        InternationalMacroSheetRelationship,
        InternationalMacroSheetContentType,
        "macrosheet",
        MacroSheetNamespace,
        SpreadsheetNamespace)]
    [InlineData(
        DialogSheetRelationship,
        DialogSheetContentType,
        "dialogsheet",
        SpreadsheetNamespace,
        SpreadsheetNamespace)]
    [InlineData(
        StrictDialogSheetRelationship,
        DialogSheetContentType,
        "dialogsheet",
        StrictSpreadsheetNamespace,
        StrictSpreadsheetNamespace)]
    public void HasValidWorkbookContainer_ValidExcelSheetType_ReturnsTrue(
        string relationshipType,
        string contentType,
        string sheetRoot,
        string sheetNamespace,
        string workbookNamespace)
    {
        var filePath = CreateWorkbookPackage(
            relationshipType,
            contentType,
            sheetRoot,
            sheetNamespace,
            workbookNamespace);

        Assert.True(FileAccessValidator.HasValidWorkbookContainer(filePath));
    }

    [Fact]
    public void HasValidWorkbookContainer_MacroSheetWithWorksheetRelationship_ReturnsFalse()
    {
        var filePath = CreateWorkbookPackage(
            WorksheetRelationship,
            MacroSheetContentType,
            "macrosheet",
            MacroSheetNamespace,
            SpreadsheetNamespace);

        Assert.False(FileAccessValidator.HasValidWorkbookContainer(filePath));
    }

    [Fact]
    public void HasValidWorkbookContainer_DialogSheetWithWorksheetContentType_ReturnsFalse()
    {
        var filePath = CreateWorkbookPackage(
            DialogSheetRelationship,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            "dialogsheet",
            SpreadsheetNamespace,
            SpreadsheetNamespace);

        Assert.False(FileAccessValidator.HasValidWorkbookContainer(filePath));
    }

    [Fact]
    public void HasValidWorkbookContainer_MacroSheetWithSpreadsheetRootNamespace_ReturnsFalse()
    {
        var filePath = CreateWorkbookPackage(
            MacroSheetRelationship,
            MacroSheetContentType,
            "macrosheet",
            SpreadsheetNamespace,
            SpreadsheetNamespace);

        Assert.False(FileAccessValidator.HasValidWorkbookContainer(filePath));
    }

    [Fact]
    public void IsIrmProtected_PasswordEncryptionMarkersWithoutDrmDataSpace_ReturnsFalse()
    {
        var filePath = fixture.CreateFilePath();
        var bytes = new byte[2048];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
            .CopyTo(bytes, 0);
        System.Text.Encoding.Unicode.GetBytes("EncryptionInfo")
            .CopyTo(bytes, 512);
        System.Text.Encoding.Unicode.GetBytes("EncryptedPackage")
            .CopyTo(bytes, 1024);
        File.WriteAllBytes(filePath, bytes);

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_LegacyDrmDataSpaceMetadata_ReturnsTrue()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "\tDRMDataSpace");

        Assert.True(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_ModernDrmEncryptedDataSpaceMetadata_ReturnsTrue()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "DRMEncryptedDataSpace");

        Assert.True(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_SimilarDataSpaceName_ReturnsFalse()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "DRMEncryptedDataSpacePreview");

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_MapWithoutMatchingDefinition_ReturnsFalse()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "\tDRMDataSpace",
            "DifferentDataSpace");

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_DeepDirectoryTree_ReturnsFalseWithoutRecursionFailure()
    {
        var filePath = OleDataSpaceTestFile.WriteDeepDirectory(
            fixture.CreateFilePath(),
            entryCount: 10_000);

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_ThousandsOfNestedMaximumLengthStorages_HasBoundedAllocation()
    {
        var filePath = OleDataSpaceTestFile.WriteDeepNestedStorages(
            fixture.CreateFilePath(),
            entryCount: 2_000);
        var before = GC.GetAllocatedBytesForCurrentThread();

        var isIrmProtected = FileAccessValidator.IsIrmProtected(filePath);

        var allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.False(isIrmProtected);
        Assert.True(
            allocatedBytes < 32 * 1024 * 1024,
            $"Nested OLE inspection allocated {allocatedBytes:N0} bytes.");
    }

    [Fact]
    public void IsIrmProtected_OversizedVersion4RootMiniStream_ReturnsFalse()
    {
        var filePath = OleDataSpaceTestFile.WriteOversizedVersion4RootMiniStream(
            fixture.CreateFilePath());

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void GetSectorCount_SparseFileBeyondInspectionLimit_ThrowsInvalidDataException()
    {
        var oversizedLength = checked(((long)int.MaxValue + 2) * 512);

        var exception = Assert.Throws<InvalidDataException>(
            () => OleCompoundFileReader.GetSectorCount(oversizedLength, 512));

        Assert.Contains("length", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private string CreateWorkbookPackage(
        string relationshipType,
        string sheetContentType,
        string sheetRoot,
        string sheetNamespace,
        string workbookNamespace)
    {
        const string contentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        const string packageRelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        var strict = workbookNamespace == StrictSpreadsheetNamespace;
        var officeRelationshipNamespace = strict
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var officeDocumentRelationship =
            $"{officeRelationshipNamespace}/officeDocument";
        var sheetPath = sheetRoot switch
        {
            "dialogsheet" => "dialogsheets/sheet1.xml",
            "macrosheet" => "macrosheets/sheet1.xml",
            "chartsheet" => "chartsheets/sheet1.xml",
            _ => "worksheets/sheet1.xml"
        };
        var filePath = fixture.CreateFilePath();

        using var archive = ZipFile.Open(filePath, ZipArchiveMode.Create);
        WriteEntry(
            archive,
            "[Content_Types].xml",
            $"""
             <Types xmlns="{contentTypesNamespace}">
               <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
               <Override PartName="/xl/{sheetPath}" ContentType="{sheetContentType}"/>
             </Types>
             """);
        WriteEntry(
            archive,
            "_rels/.rels",
            $"""
             <Relationships xmlns="{packageRelationshipsNamespace}">
               <Relationship Id="rId1" Type="{officeDocumentRelationship}" Target="xl/workbook.xml"/>
             </Relationships>
             """);
        WriteEntry(
            archive,
            "xl/workbook.xml",
            $"""
             <workbook xmlns="{workbookNamespace}" xmlns:r="{officeRelationshipNamespace}">
               <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
             </workbook>
             """);
        WriteEntry(
            archive,
            "xl/_rels/workbook.xml.rels",
            $"""
             <Relationships xmlns="{packageRelationshipsNamespace}">
               <Relationship Id="rId1" Type="{relationshipType}" Target="{sheetPath}"/>
             </Relationships>
             """);
        WriteEntry(
            archive,
            $"xl/{sheetPath}",
            $"""<{sheetRoot} xmlns="{sheetNamespace}"/>""");

        return filePath;
    }

    private static void WriteEntry(
        ZipArchive archive,
        string path,
        string content)
    {
        var entry = archive.CreateEntry(path);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }
}
