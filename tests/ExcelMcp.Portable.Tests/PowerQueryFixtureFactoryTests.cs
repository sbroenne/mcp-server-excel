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
