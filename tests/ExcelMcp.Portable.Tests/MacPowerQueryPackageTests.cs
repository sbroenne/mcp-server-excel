using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryPackageTests
{
    [Fact]
    public void ReadQueries_WorkbookWithoutDataMashup_ReturnsEmptyList()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-pq-empty-");
        var path = Path.Combine(directory.FullName, "empty.xlsx");
        using (ZipFile.Open(path, ZipArchiveMode.Create))
        {
        }

        try
        {
            Assert.Empty(MacPowerQueryPackage.ReadQueries(path));
        }
        finally
        {
            File.Delete(path);
            directory.Delete();
        }
    }

    [Fact]
    public void ReadQueries_ParsesPlainAndQuotedIdentifiersWithoutSplittingStrings()
    {
        var path = CreateWorkbook(
            """
            section Section1;

            shared Sales = let
                Source = #table({"Text"}, {{"semi;colon"}})
            in
                Source;

            shared #"Monthly Sales" = let
                Source = Sales
            in
                Source;
            """);

        try
        {
            var queries = MacPowerQueryPackage.ReadQueries(path);

            Assert.Equal(["Sales", "Monthly Sales"], queries.Select(query => query.Name));
            Assert.Contains("""{{"semi;colon"}}""", queries[0].Formula, StringComparison.Ordinal);
            Assert.Contains("Source = Sales", queries[1].Formula, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(Path.GetDirectoryName(path)!);
        }
    }

    [Fact]
    public void UpdateQuery_RewritesOnlyExactQueryAndResetsPermissions()
    {
        var path = CreateWorkbook(
            """
            section Section1;
            shared Sales = 1;
            shared SalesArchive = 2;
            """);

        try
        {
            MacPowerQueryPackage.UpdateQuery(path, "Sales", """#table({"Value"}, {{42}})""");

            var queries = MacPowerQueryPackage.ReadQueries(path);
            Assert.Equal("""#table({"Value"}, {{42}})""", queries[0].Formula);
            Assert.Equal("2", queries[1].Formula);
            Assert.Contains(
                "<FirewallEnabled>true</FirewallEnabled>",
                ReadPermissions(path),
                StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(Path.GetDirectoryName(path)!);
        }
    }

    [Fact]
    public void UpdateQuery_MissingExactName_DoesNotModifyWorkbook()
    {
        var path = CreateWorkbook("section Section1; shared SalesArchive = 2;");
        var before = File.ReadAllBytes(path);

        try
        {
            var error = Assert.Throws<InvalidOperationException>(
                () => MacPowerQueryPackage.UpdateQuery(path, "Sales", "3"));

            Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(Path.GetDirectoryName(path)!);
        }
    }

    [Fact]
    public void ReadWorksheetLoads_FollowsWorksheetTableAndQueryTableRelationships()
    {
        var path = CreateWorkbook("section Section1; shared Sales = 1;");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            WriteEntry(
                archive,
                "xl/workbook.xml",
                Encoding.UTF8.GetBytes(
                    """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>"""));
            WriteEntry(
                archive,
                "xl/_rels/workbook.xml.rels",
                Encoding.UTF8.GetBytes(
                    """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""));
            WriteEntry(
                archive,
                "xl/worksheets/_rels/sheet1.xml.rels",
                Encoding.UTF8.GetBytes(
                    """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"""));
            WriteEntry(
                archive,
                "xl/tables/_rels/table1.xml.rels",
                Encoding.UTF8.GetBytes(
                    """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>"""));
            WriteEntry(
                archive,
                "xl/queryTables/queryTable1.xml",
                Encoding.UTF8.GetBytes(
                    """<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" connectionId="7"/>"""));
            WriteEntry(
                archive,
                "xl/connections.xml",
                Encoding.UTF8.GetBytes(
                    """<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="7"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Sales"/></connection></connections>"""));
        }

        try
        {
            var load = Assert.Single(MacPowerQueryPackage.ReadWorksheetLoads(path));

            Assert.Equal("Report", load.SheetName);
            Assert.Contains("Location=Sales", load.Connection, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(Path.GetDirectoryName(path)!);
        }
    }

    [Fact]
    public void ReadWorksheetLoads_WorkbookWithDataModel_RejectsAmbiguousLoadState()
    {
        var path = CreateWorkbook("section Section1; shared Sales = 1;");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            WriteEntry(archive, "xl/model/model.bin", [1, 2, 3]);
        }

        try
        {
            var error = Assert.Throws<NotSupportedException>(
                () => MacPowerQueryPackage.ReadWorksheetLoads(path));

            Assert.Contains("Data Model", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(Path.GetDirectoryName(path)!);
        }
    }

    private static string CreateWorkbook(string section)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-pq-package-");
        var path = Path.Combine(directory.FullName, "query.xlsx");
        var packageParts = CreateZip(
            ("Formulas/Section1.m", Encoding.UTF8.GetBytes(section)),
            ("Config/Package.xml", Encoding.UTF8.GetBytes("<Package />")));
        var metadata = BuildMetadata(CreateZip());
        var root = BuildRoot(
            packageParts,
            Encoding.UTF8.GetBytes("<PermissionList><FirewallEnabled>false</FirewallEnabled></PermissionList>"),
            metadata,
            []);
        var xml = $"""<?xml version="1.0" encoding="utf-8"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">{Convert.ToBase64String(root)}</DataMashup>""";

        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteEntry(archive, "customXml/item1.xml", Encoding.UTF8.GetBytes(xml));
        return path;
    }

    private static string ReadPermissions(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = Assert.Single(
            archive.Entries,
            candidate => candidate.FullName == "customXml/item1.xml");
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        var xml = XDocument.Parse(reader.ReadToEnd());
        var root = Convert.FromBase64String(xml.Root!.Value);
        using var stream = new MemoryStream(root);
        using var binary = new BinaryReader(stream, Encoding.UTF8);
        _ = binary.ReadUInt32();
        _ = binary.ReadBytes(binary.ReadInt32());
        return Encoding.UTF8.GetString(binary.ReadBytes(binary.ReadInt32()));
    }

    private static byte[] BuildMetadata(byte[] content)
    {
        var xml = Encoding.UTF8.GetBytes("<LocalPackageMetadataFile />");
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(0u);
        writer.Write(xml.Length);
        writer.Write(xml);
        writer.Write(content.Length);
        writer.Write(content);
        return stream.ToArray();
    }

    private static byte[] BuildRoot(
        byte[] packageParts,
        byte[] permissions,
        byte[] metadata,
        byte[] permissionBindings)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(0u);
        WriteSized(writer, packageParts);
        WriteSized(writer, permissions);
        WriteSized(writer, metadata);
        WriteSized(writer, permissionBindings);
        return stream.ToArray();
    }

    private static byte[] CreateZip(params (string Path, byte[] Content)[] entries)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (path, content) in entries)
            {
                WriteEntry(archive, path, content);
            }
        }
        return stream.ToArray();
    }

    private static void WriteSized(BinaryWriter writer, byte[] value)
    {
        writer.Write(value.Length);
        writer.Write(value);
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] content)
    {
        using var target = archive.CreateEntry(path, CompressionLevel.Optimal).Open();
        target.Write(content);
    }
}
