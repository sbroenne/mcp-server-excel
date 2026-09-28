using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class PowerQueryFixtureFactoryTests
{
    [Theory]
    [InlineData(PowerQueryFixtureKind.ConnectionOnly)]
    [InlineData(PowerQueryFixtureKind.WorksheetLoaded)]
    public void Create_ProducesAuditedRepositoryOwnedLiteralQueryPackage(
        PowerQueryFixtureKind kind)
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(directory.Path, kind);

        var audit = PowerQueryFixtureFactory.Audit(fixture.WorkbookPath, kind);
        Assert.Empty(audit.Errors);
        Assert.Contains("customXml/item1.xml", audit.PackageParts);
        Assert.Contains("xl/connections.xml", audit.PackageParts);
        Assert.Equal(PowerQueryFixtureFactory.LiteralM, Assert.Single(
            MacPowerQueryPackage.ReadQueries(fixture.WorkbookPath)).Formula);

        var manifest = JsonSerializer.Deserialize<PowerQueryFixtureManifest>(
            File.ReadAllText(fixture.ProvenanceManifestPath));
        Assert.NotNull(manifest);
        Assert.Equal("repository-authored", manifest.Provenance);
        Assert.Equal(PowerQueryFixtureFactory.QueryName, manifest.QueryName);
        Assert.Equal(kind.ToString(), manifest.LoadKind);
        Assert.Contains("MS-QDEFF", manifest.Specifications);
        Assert.Contains("ECMA-376", manifest.Specifications);
        Assert.Equal(audit.Sha256, manifest.WorkbookSha256);
    }

    [Fact]
    public void WorksheetLoadedFixture_HasExactWorksheetTableQueryTableConnectionGraph()
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(
            directory.Path,
            PowerQueryFixtureKind.WorksheetLoaded);

        var load = Assert.Single(
            MacPowerQueryPackage.ReadWorksheetLoads(fixture.WorkbookPath));
        Assert.Equal(PowerQueryFixtureFactory.WorksheetName, load.SheetName);
        Assert.Contains(
            $"Location={PowerQueryFixtureFactory.QueryName}",
            load.Connection,
            StringComparison.Ordinal);
        using var archive = ZipFile.OpenRead(fixture.WorkbookPath);
        using var stream = archive.GetEntry("xl/tables/table1.xml")!.Open();
        var table = XDocument.Load(stream);
        Assert.Equal("queryTable", (string?)table.Root?.Attribute("tableType"));
    }

    [Fact]
    public void ConnectionOnlyFixture_HasNoWorksheetLoadGraph()
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(
            directory.Path,
            PowerQueryFixtureKind.ConnectionOnly);

        Assert.Empty(MacPowerQueryPackage.ReadWorksheetLoads(fixture.WorkbookPath));
    }

    [Theory]
    [InlineData(PowerQueryFixtureKind.ConnectionOnly)]
    [InlineData(PowerQueryFixtureKind.WorksheetLoaded)]
    public void Create_IsByteDeterministic(PowerQueryFixtureKind kind)
    {
        using var firstDirectory = new TemporaryDirectory("excelmcp-pq-fixture-");
        using var secondDirectory = new TemporaryDirectory("excelmcp-pq-fixture-");

        var first = PowerQueryFixtureFactory.Create(firstDirectory.Path, kind);
        var second = PowerQueryFixtureFactory.Create(secondDirectory.Path, kind);

        Assert.Equal(
            File.ReadAllBytes(first.WorkbookPath),
            File.ReadAllBytes(second.WorkbookPath));
    }

    [Theory]
    [InlineData(PowerQueryFixtureKind.ConnectionOnly)]
    [InlineData(PowerQueryFixtureKind.WorksheetLoaded)]
    public void Create_EmitsThreeStylesInEachThemeMatrixList(PowerQueryFixtureKind kind)
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(directory.Path, kind);
        using var archive = ZipFile.OpenRead(fixture.WorkbookPath);
        using var stream = archive.GetEntry("xl/theme/theme1.xml")!.Open();
        var theme = XDocument.Load(stream);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var matrix = Assert.Single(theme.Descendants(drawing + "fmtScheme"));
        Assert.Equal(4, matrix.Elements().Count());
        Assert.All(matrix.Elements(), list => Assert.Equal(3, list.Elements().Count()));
    }

    [Fact]
    public void Audit_RejectsIncompleteThemeMatrixInsteadOfClaimingPackageValidity()
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(directory.Path, PowerQueryFixtureKind.ConnectionOnly);
        using (var archive = ZipFile.Open(fixture.WorkbookPath, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/theme/theme1.xml")!;
            XDocument theme;
            using (var stream = entry.Open()) theme = XDocument.Load(stream);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            foreach (var list in theme.Descendants(drawing + "fmtScheme").Elements())
            {
                list.Elements().Skip(1).Remove();
            }
            entry.Delete();
            using var output = archive.CreateEntry("xl/theme/theme1.xml").Open();
            theme.Save(output);
        }

        var audit = PowerQueryFixtureFactory.Audit(fixture.WorkbookPath, PowerQueryFixtureKind.ConnectionOnly);
        foreach (var name in new[] { "fillStyleLst", "lnStyleLst", "effectStyleLst", "bgFillStyleLst" })
        {
            Assert.Contains(audit.Errors, error => error.Contains(name, StringComparison.Ordinal));
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("worksheet")]
    [InlineData("xml")]
    public void Audit_RejectsTableTypeThatDoesNotMatchQueryTableRelationship(string? tableType)
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(directory.Path, PowerQueryFixtureKind.WorksheetLoaded);
        using (var archive = ZipFile.Open(fixture.WorkbookPath, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/tables/table1.xml")!;
            XDocument table;
            using (var stream = entry.Open()) table = XDocument.Load(stream);
            table.Root!.SetAttributeValue("tableType", tableType);
            entry.Delete();
            using var output = archive.CreateEntry("xl/tables/table1.xml").Open();
            table.Save(output);
        }

        var audit = PowerQueryFixtureFactory.Audit(fixture.WorkbookPath, PowerQueryFixtureKind.WorksheetLoaded);
        Assert.Contains(audit.Errors, error => error.Contains("tableType", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(PowerQueryFixtureKind.ConnectionOnly)]
    [InlineData(PowerQueryFixtureKind.WorksheetLoaded)]
    public void Create_EncodesMashupColumnNamesAsJsonArray(PowerQueryFixtureKind kind)
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(directory.Path, kind);
        RewriteMetadata(fixture.WorkbookPath, metadata =>
        {
            var entry = Assert.Single(metadata.Descendants(),
                element => (string?)element.Attribute("Type") == "FillColumnNames");
            var value = Assert.IsType<string>((string?)entry.Attribute("Value"));
            Assert.StartsWith("s", value, StringComparison.Ordinal);
            var names = JsonSerializer.Deserialize<string[]>(value[1..]);
            Assert.NotNull(names);
            Assert.Equal(["Item", "Amount"], names);
        });
    }

    [Theory]
    [InlineData("sItem,Amount")]
    [InlineData("s[\"Item\"]")]
    [InlineData("l2")]
    public void Audit_RejectsInvalidMashupColumnNames(string value)
    {
        using var directory = new TemporaryDirectory("excelmcp-pq-fixture-");
        var fixture = PowerQueryFixtureFactory.Create(directory.Path, PowerQueryFixtureKind.WorksheetLoaded);
        RewriteMetadata(fixture.WorkbookPath, metadata =>
            metadata.Descendants().Single(element =>
                (string?)element.Attribute("Type") == "FillColumnNames").SetAttributeValue("Value", value));

        var audit = PowerQueryFixtureFactory.Audit(fixture.WorkbookPath, PowerQueryFixtureKind.WorksheetLoaded);
        Assert.Contains(audit.Errors, error => error.Contains("FillColumnNames", StringComparison.Ordinal));
    }

    private static void RewriteMetadata(string path, Action<XDocument> edit)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry("customXml/item1.xml")!;
        XDocument document;
        using (var stream = entry.Open()) document = XDocument.Load(stream);
        using var rootReader = new BinaryReader(new MemoryStream(Convert.FromBase64String(document.Root!.Value)));
        var rootVersion = rootReader.ReadUInt32();
        var sections = Enumerable.Range(0, 4)
            .Select(_ => rootReader.ReadBytes(rootReader.ReadInt32())).ToArray();
        using var metadataReader = new BinaryReader(new MemoryStream(sections[2]));
        var metadataVersion = metadataReader.ReadUInt32();
        using var xmlStream = new MemoryStream(metadataReader.ReadBytes(metadataReader.ReadInt32()));
        var metadata = XDocument.Load(xmlStream);
        var binaryContent = metadataReader.ReadBytes(metadataReader.ReadInt32());
        edit(metadata);
        using var updatedXml = new MemoryStream();
        metadata.Save(updatedXml);
        using var updatedMetadata = new MemoryStream();
        using (var writer = new BinaryWriter(updatedMetadata, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            writer.Write(metadataVersion);
            writer.Write(checked((int)updatedXml.Length));
            writer.Write(updatedXml.ToArray());
            writer.Write(binaryContent.Length);
            writer.Write(binaryContent);
        }
        sections[2] = updatedMetadata.ToArray();
        using var updatedRoot = new MemoryStream();
        using (var writer = new BinaryWriter(updatedRoot, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            writer.Write(rootVersion);
            foreach (var section in sections)
            {
                writer.Write(section.Length);
                writer.Write(section);
            }
        }
        document.Root.Value = Convert.ToBase64String(updatedRoot.ToArray());
        entry.Delete();
        using var output = archive.CreateEntry("customXml/item1.xml").Open();
        document.Save(output);
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory(string prefix)
        {
            Path = Directory.CreateTempSubdirectory(prefix).FullName;
        }

        public string Path { get; }

        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
