using System.Xml.Linq;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.XmlMap;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.XmlMap;

/// <summary>
/// Integration tests for XML map lifecycle and in-memory XML import/export.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "XmlMap")]
[Trait("RequiresExcel", "true")]
public sealed class XmlMapCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private const string CustomerSchema = """
        <?xml version="1.0" encoding="utf-8"?>
        <xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
          <xs:element name="customer">
            <xs:complexType>
              <xs:sequence>
                <xs:element name="name" type="xs:string" />
              </xs:sequence>
            </xs:complexType>
          </xs:element>
        </xs:schema>
        """;

    private readonly XmlMapCommands _commands = new();
    private readonly TempDirectoryFixture _fixture;

    public XmlMapCommandsTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddListDelete_RoundTripsXmlMapLifecycle()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        var addResult = _commands.Add(batch, CustomerSchema, "customer", "CustomerMap");
        Assert.True(addResult.Success);
        Assert.Equal("CustomerMap", addResult.MapName);

        var listResult = _commands.List(batch);
        var map = Assert.Single(listResult.Maps);
        Assert.Equal("CustomerMap", map.Name);
        Assert.Equal("customer", map.RootElementName);

        var deleteResult = _commands.Delete(batch, "CustomerMap");
        Assert.True(deleteResult.Success);
        Assert.Empty(_commands.List(batch).Maps);
    }

    [Fact]
    public void MapRangeImportExport_RoundTripsMappedCell()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        _commands.Add(batch, CustomerSchema, "customer", "CustomerMap");
        var mapResult = _commands.MapRange(
            batch,
            "CustomerMap",
            "Sheet1",
            "A1",
            "/customer/name");
        Assert.True(mapResult.Success);

        var importResult = _commands.ImportXml(
            batch,
            "<customer><name>Ada Lovelace</name></customer>",
            mapName: "CustomerMap");
        Assert.True(importResult.Success);
        Assert.Equal("CustomerMap", importResult.MapName);

        var exportResult = _commands.ExportXml(batch, "CustomerMap");
        Assert.True(exportResult.Success);
        Assert.Equal("Ada Lovelace", XDocument.Parse(exportResult.XmlData).Root?.Element("name")?.Value);
    }

    [Fact]
    public void ImportXml_WithDestination_CreatesMapAndExportsImportedData()
    {
        const string xmlData = """
            <customers>
              <customer><name>Ada</name><score>42</score></customer>
              <customer><name>Grace</name><score>99</score></customer>
            </customers>
            """;
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        var importResult = _commands.ImportXml(
            batch,
            xmlData,
            sheetName: "Sheet1",
            startCell: "B2");

        Assert.True(importResult.Success);
        Assert.False(string.IsNullOrWhiteSpace(importResult.MapName));
        Assert.Equal("Sheet1", importResult.SheetName);
        Assert.Equal("B2", importResult.StartCell);

        var exportResult = _commands.ExportXml(batch, importResult.MapName);
        Assert.True(exportResult.Success);
        var exported = XDocument.Parse(exportResult.XmlData).ToString(SaveOptions.DisableFormatting);
        Assert.Contains("Ada", exported, StringComparison.Ordinal);
        Assert.Contains("Grace", exported, StringComparison.Ordinal);
        Assert.Contains("42", exported, StringComparison.Ordinal);
        Assert.Contains("99", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void ImportXml_WithExistingMap_OverwritesMappedCell()
    {
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        _commands.Add(batch, CustomerSchema, "customer", "CustomerMap");
        _commands.MapRange(batch, "CustomerMap", "Sheet1", "A1", "/customer/name");
        _commands.ImportXml(batch, "<customer><name>First</name></customer>", mapName: "CustomerMap");
        _commands.ImportXml(batch, "<customer><name>Second</name></customer>", mapName: "CustomerMap");

        var exportResult = _commands.ExportXml(batch, "CustomerMap");
        Assert.Equal("Second", XDocument.Parse(exportResult.XmlData).Root?.Element("name")?.Value);
        Assert.DoesNotContain("First", exportResult.XmlData, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("urn:blocked")]
    [InlineData("https://example.invalid/schema.xsd")]
    [InlineData("\\\\127.0.0.1\\missing\\schema.xsd")]
    [InlineData("file:///C:/nonexistent/schema.xsd")]
    public void ImportXml_WithSchemaLocation_IsRejectedBeforeAutomaticMapping(string schemaLocation)
    {
        var xmlData = $"""
            <customers
                xmlns="urn:customers"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:schemaLocation="urn:customers {schemaLocation}">
              <customer><name>Ada</name></customer>
            </customers>
            """;
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.ImportXml(batch, xmlData, sheetName: "Sheet1"));

        Assert.Contains("schema location", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(_commands.List(batch).Maps);
    }

    [Theory]
    [InlineData("urn:blocked")]
    [InlineData("https://example.invalid/schema.xsd")]
    [InlineData("\\\\127.0.0.1\\missing\\schema.xsd")]
    [InlineData("file:///C:/nonexistent/schema.xsd")]
    public void ImportXml_WithNoNamespaceSchemaLocation_IsRejectedBeforeAutomaticMapping(string schemaLocation)
    {
        var xmlData = $"""
            <customers
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:noNamespaceSchemaLocation="{schemaLocation}">
              <customer><name>Ada</name></customer>
            </customers>
            """;
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.ImportXml(batch, xmlData, sheetName: "Sheet1"));

        Assert.Contains("schema location", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(_commands.List(batch).Maps);
    }

    [Fact]
    public void Add_SchemaWithExternalDependency_IsRejected()
    {
        const string schemaWithImport = """
            <xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
              <xs:import namespace="urn:external" schemaLocation="https://example.com/external.xsd" />
              <xs:element name="root" type="xs:string" />
            </xs:schema>
            """;
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.Add(batch, schemaWithImport, "root", "UnsafeMap"));

        Assert.Contains("external schema", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(_commands.List(batch).Maps);
    }

    [Fact]
    public void Add_SchemaWithExternalRedefine_IsRejected()
    {
        const string schemaWithRedefine = """
            <xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
              <xs:redefine schemaLocation="https://example.com/base.xsd">
                <xs:complexType name="ExternalType">
                  <xs:complexContent>
                    <xs:extension base="ExternalType" />
                  </xs:complexContent>
                </xs:complexType>
              </xs:redefine>
              <xs:element name="root" type="ExternalType" />
            </xs:schema>
            """;
        var filePath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(filePath);

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.Add(batch, schemaWithRedefine, "root", "UnsafeMap"));

        Assert.Contains("external schema", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(_commands.List(batch).Maps);
    }
}
