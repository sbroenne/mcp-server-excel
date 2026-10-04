using System.Xml.Linq;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceXmlMapTests
{
    [Fact]
    public void AddListDelete_RoundTripsXmlMapLifecycle()
    {
        var batch = _fixture.BatchToken;

        var guard = CreateGuard();
        var before = CaptureGuard(guard);
        var addResult = RequireSuccess(_xmlMapCommands.Add(batch, CustomerSchema, "customer", "CustomerMap"));
        _fixture.RegisterXmlMapForCleanup("CustomerMap");
        Assert.True(addResult.Success);
        Assert.Equal("CustomerMap", addResult.MapName);

        var listResult = RequireSuccess(_xmlMapCommands.List(batch));
        Assert.Equal(2, listResult.Maps.Count);
        var map = Assert.Single(listResult.Maps, item => item.Name == "CustomerMap");
        Assert.Equal("CustomerMap", map.Name);
        Assert.Equal("customer", map.RootElementName);

        RequireSuccess(_xmlMapCommands.MapRange(batch, "CustomerMap", guard.Sheet, "D3", "/customer/name"));
        RequireSuccess(_xmlMapCommands.ImportXml(batch, "<customer><name>Deleted map cells remain</name></customer>",
            mapName: "CustomerMap"));
        Assert.Equal("Deleted map cells remain",
            Assert.Single(Assert.Single(RequireSuccess(_commands.GetValues(batch, guard.Sheet, "D3")).Values)));
        var deleteResult = RequireSuccess(_xmlMapCommands.Delete(batch, "CustomerMap"));
        _fixture.ForgetXmlMap("CustomerMap");
        Assert.True(deleteResult.Success);
        Assert.Equal("GuardMap", Assert.Single(RequireSuccess(_xmlMapCommands.List(batch)).Maps).Name);
        Assert.Equal("Deleted map cells remain",
            Assert.Single(Assert.Single(RequireSuccess(_commands.GetValues(batch, guard.Sheet, "D3")).Values)));
        Assert.Equal(before, CaptureGuard(guard));
    }

    [Fact]
    public void MapRangeImportExport_RoundTripsMappedCell()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "D1:E2", [["Unmapped", 31], ["Retained", 59]]));

        RequireSuccess(_xmlMapCommands.Add(batch, CustomerSchema, "customer", "CustomerMap"));
        _fixture.RegisterXmlMapForCleanup("CustomerMap");
        var mapResult = RequireSuccess(_xmlMapCommands.MapRange(
            batch,
            "CustomerMap",
            sheet,
            "A1",
            "/customer/name"));
        Assert.True(mapResult.Success);
        AssertMapping(sheet, "A1", "CustomerMap", "/customer/name", repeating: false);

        var importResult = RequireSuccess(_xmlMapCommands.ImportXml(
            batch,
            "<customer><name>Ada Lovelace</name></customer>",
            mapName: "CustomerMap"));
        Assert.True(importResult.Success);
        Assert.Equal("CustomerMap", importResult.MapName);
        Assert.Equal("xlXmlImportSuccess", importResult.ImportStatus);
        Assert.Null(importResult.SheetName);
        Assert.Null(importResult.StartCell);
        var mapped = RequireSuccess(_commands.GetValues(batch, sheet, "A1"));
        Assert.True(mapped.Success, mapped.ErrorMessage);
        Assert.Equal("Ada Lovelace", mapped.Values[0][0]);

        var exportResult = RequireSuccess(_xmlMapCommands.ExportXml(batch, "CustomerMap"));
        Assert.True(exportResult.Success);
        AssertCustomer(exportResult.XmlData, "Ada Lovelace");
        Assert.Equal("CustomerMap", exportResult.MapName);
        Assert.Equal("xlXmlExportSuccess", exportResult.ExportStatus);
        PowerQueryStateAssertions.AssertRows([["Unmapped", 31], ["Retained", 59]],
            RequireSuccess(_commands.GetValues(batch, sheet, "D1:E2")).Values);
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
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "F2:G3", [["Neighbor", 17], ["Retained", 29]]));

        var importResult = RequireSuccess(_xmlMapCommands.ImportXml(
            batch,
            xmlData,
            sheetName: sheet,
            startCell: "B2"));
        _fixture.RegisterXmlMapForCleanup(importResult.MapName);

        Assert.True(importResult.Success);
        Assert.False(string.IsNullOrWhiteSpace(importResult.MapName));
        Assert.Equal(sheet, importResult.SheetName);
        Assert.Equal("B2", importResult.StartCell);
        Assert.Equal("xlXmlImportSuccess", importResult.ImportStatus);

        var exportResult = RequireSuccess(_xmlMapCommands.ExportXml(batch, importResult.MapName));
        Assert.True(exportResult.Success);
        var root = XDocument.Parse(exportResult.XmlData).Root;
        Assert.NotNull(root);
        Assert.Equal("customers", root.Name.LocalName);
        Assert.Equal(
            [("Ada", "42"), ("Grace", "99")],
            root.Elements("customer").Select(customer =>
                (customer.Element("name")?.Value, customer.Element("score")?.Value)));
        Assert.Equal(importResult.MapName, exportResult.MapName);
        Assert.Equal("xlXmlExportSuccess", exportResult.ExportStatus);
        var loaded = RequireSuccess(_commands.GetValues(batch, sheet, "B2:C4"));
        Assert.True(loaded.Success, loaded.ErrorMessage);
        PowerQueryStateAssertions.AssertRows([["name", "score"], ["Ada", 42], ["Grace", 99]], loaded.Values);
        AssertMapping(sheet, "B3", importResult.MapName, "/customers/customer/name", repeating: true);
        AssertMapping(sheet, "C3", importResult.MapName, "/customers/customer/score", repeating: true);
        PowerQueryStateAssertions.AssertRows([["Neighbor", 17], ["Retained", 29]],
            RequireSuccess(_commands.GetValues(batch, sheet, "F2:G3")).Values);
    }

    [Fact]
    public void ImportXml_WithExistingMap_OverwritesMappedCell()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "D1:E2", [["Unmapped", 31], ["Retained", 59]]));

        RequireSuccess(_xmlMapCommands.Add(batch, CustomerSchema, "customer", "CustomerMap"));
        _fixture.RegisterXmlMapForCleanup("CustomerMap");
        RequireSuccess(_xmlMapCommands.MapRange(batch, "CustomerMap", sheet, "A1", "/customer/name"));
        AssertMapping(sheet, "A1", "CustomerMap", "/customer/name", repeating: false);
        var first = RequireSuccess(_xmlMapCommands.ImportXml(batch, "<customer><name>First</name></customer>", mapName: "CustomerMap"));
        Assert.True(first.Success, first.ErrorMessage);
        Assert.Equal("First", RequireSuccess(_commands.GetValues(batch, sheet, "A1")).Values[0][0]);
        AssertCustomer(RequireSuccess(_xmlMapCommands.ExportXml(batch, "CustomerMap")).XmlData, "First");
        RequireSuccess(_xmlMapCommands.ImportXml(batch, "<customer><name>Second</name></customer>", mapName: "CustomerMap"));

        var exportResult = RequireSuccess(_xmlMapCommands.ExportXml(batch, "CustomerMap"));
        Assert.True(exportResult.Success, exportResult.ErrorMessage);
        AssertCustomer(exportResult.XmlData, "Second");
        Assert.DoesNotContain("First", exportResult.XmlData, StringComparison.Ordinal);
        var replaced = RequireSuccess(_commands.GetValues(batch, sheet, "A1"));
        Assert.True(replaced.Success, replaced.ErrorMessage);
        Assert.Equal("Second", replaced.Values[0][0]);
        AssertMapping(sheet, "A1", "CustomerMap", "/customer/name", repeating: false);
        PowerQueryStateAssertions.AssertRows([["Unmapped", 31], ["Retained", 59]],
            RequireSuccess(_commands.GetValues(batch, sheet, "D1:E2")).Values);
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
        var batch = _fixture.BatchToken;
        var guard = CreateGuard();
        var before = CaptureGuard(guard);

        var exception = Assert.Throws<ArgumentException>(
            () => _xmlMapCommands.ImportXml(batch, xmlData, sheetName: guard.Sheet));

        Assert.Contains("schema location", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, CaptureGuard(guard));
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
        var batch = _fixture.BatchToken;
        var guard = CreateGuard();
        var before = CaptureGuard(guard);

        var exception = Assert.Throws<ArgumentException>(
            () => _xmlMapCommands.ImportXml(batch, xmlData, sheetName: guard.Sheet));

        Assert.Contains("schema location", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, CaptureGuard(guard));
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
        var batch = _fixture.BatchToken;
        var guard = CreateGuard();
        var before = CaptureGuard(guard);

        var exception = Assert.Throws<ArgumentException>(
            () => _xmlMapCommands.Add(batch, schemaWithImport, "root", "UnsafeMap"));

        Assert.Contains("external schema", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, CaptureGuard(guard));
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
        var batch = _fixture.BatchToken;
        var guard = CreateGuard();
        var before = CaptureGuard(guard);

        var exception = Assert.Throws<ArgumentException>(
            () => _xmlMapCommands.Add(batch, schemaWithRedefine, "root", "UnsafeMap"));

        Assert.Contains("external schema", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, CaptureGuard(guard));
    }

    private sealed record MappedGuard(string Sheet);

    private MappedGuard CreateGuard()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B2",
            [["Retained", 41], ["Neighbor", 67]]));
        RequireSuccess(_xmlMapCommands.Add(batch, CustomerSchema, "customer", "GuardMap"));
        _fixture.RegisterXmlMapForCleanup("GuardMap");
        RequireSuccess(_xmlMapCommands.MapRange(batch, "GuardMap", sheet, "C5", "/customer/name"));
        RequireSuccess(_xmlMapCommands.ImportXml(batch,
            "<customer><name>Mapped guard</name></customer>", mapName: "GuardMap"));
        return new(sheet);
    }

    private string CaptureGuard(MappedGuard guard)
    {
        var batch = _fixture.BatchToken;
        AssertMapping(guard.Sheet, "C5", "GuardMap", "/customer/name", repeating: false);
        var xml = RequireSuccess(_xmlMapCommands.ExportXml(batch, "GuardMap"));
        AssertCustomer(xml.XmlData, "Mapped guard");
        PowerQueryStateAssertions.AssertRows([["Retained", 41], ["Neighbor", 67]],
            RequireSuccess(_commands.GetValues(batch, guard.Sheet, "A1:B2")).Values);
        Assert.Equal("Mapped guard",
            Assert.Single(Assert.Single(RequireSuccess(_commands.GetValues(batch, guard.Sheet, "C5")).Values)));
        return JsonSerializer.Serialize(new
        {
            Maps = RequireSuccess(_xmlMapCommands.List(batch)).Maps.Where(map => map.Name == "GuardMap").ToList(),
            Export = xml,
            Cells = RequireSuccess(_commands.GetValues(batch, guard.Sheet, "A1:C2")).Values,
            NativeMaps = NativeMapNames()
        });
    }

    private static void AssertCustomer(string xml, string expected)
    {
        var root = XDocument.Parse(xml).Root;
        Assert.NotNull(root);
        Assert.Equal("customer", root.Name.LocalName);
        var name = Assert.Single(root.Elements());
        Assert.Equal("name", name.Name.LocalName);
        Assert.Equal(expected, name.Value);
        Assert.Empty(name.Elements());
    }

    private void AssertMapping(string sheetName, string address, string name, string path, bool repeating) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.XPath? xpath = null;
            Excel.XmlMap? map = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[sheetName];
                range = sheet.Range[address];
                xpath = range.XPath;
                map = xpath.Map;
                Assert.Equal(name, map.Name);
                Assert.Equal(path, xpath.Value);
                Assert.Equal(repeating, xpath.Repeating);
                Assert.True(map.IsExportable);
            }
            finally
            {
                ComUtilities.Release(ref map);
                ComUtilities.Release(ref xpath);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private string NativeMapNames() => _fixture.ExecuteRawVerification((context, _) =>
    {
        Excel.XmlMaps? maps = null;
        try
        {
            maps = context.Book.XmlMaps;
            var names = new List<string>();
            for (var index = 1; index <= maps.Count; index++)
            {
                Excel.XmlMap? map = null;
                try
                {
                    map = maps.Item[index];
                    names.Add(map.Name);
                }
                finally { ComUtilities.Release(ref map); }
            }
            return JsonSerializer.Serialize(names);
        }
        finally { ComUtilities.Release(ref maps); }
    });
}
