using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.Core.Commands.Settings;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class NativeQueryApprovalSettingsTests
{
    [Theory]
    [InlineData("l0", true, true)]
    [InlineData("l1", false, true)]
    [InlineData(null, true, false)]
    public void Read_MapsStoredValuesAndExcelDefault(string? value, bool required, bool explicitlySet)
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        CreateArchive(archivePath, SettingsXml(value));
        var settings = new NativeQueryApprovalSettings(archivePath);

        NativeQueryApprovalSnapshot result = settings.Read();

        Assert.Equal(required, result.Required);
        Assert.Equal(explicitlySet, result.ExplicitlySet);
        Assert.Equal(archivePath, settings.SettingsFilePath);
    }

    [Fact]
    public void Read_WhenSettingsEntryIsMissing_ThrowsExplicitError()
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        CreateArchive(archivePath, null);

        var exception = Assert.Throws<InvalidDataException>(() => new NativeQueryApprovalSettings(archivePath).Read());

        Assert.Contains("UserInterface/Settings.xml", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("l2")]
    [InlineData(null)]
    public void Read_WhenSettingValueIsUnsupported_Throws(string? value)
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        string xml = value is null
            ? """<UISettingsConfig><Entries><Entry Type="DisableNativeDbQueries" /></Entries></UISettingsConfig>"""
            : SettingsXml(value);
        CreateArchive(archivePath, xml);

        var exception = Assert.Throws<InvalidDataException>(() => new NativeQueryApprovalSettings(archivePath).Read());

        Assert.Contains("DisableNativeDbQueries", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Read_WhenSettingAppearsMoreThanOnce_Throws()
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        CreateArchive(
            archivePath,
            """<UISettingsConfig><Entries><Entry Type="DisableNativeDbQueries" Value="l0" /><Entry Type="DisableNativeDbQueries" Value="l1" /></Entries></UISettingsConfig>""");

        var exception = Assert.Throws<InvalidDataException>(() => new NativeQueryApprovalSettings(archivePath).Read());

        Assert.Contains("more than once", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Read_WhenArchiveIsMissing_ThrowsWithExpectedPath()
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "missing.zip");

        var exception = Assert.Throws<FileNotFoundException>(() => new NativeQueryApprovalSettings(archivePath).Read());

        Assert.Equal(archivePath, exception.FileName);
        Assert.Contains("Office 16.0", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_WhenArchiveIsMalformed_ReportsBackupAndPreservesOriginal()
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        byte[] originalContent = "not a zip archive"u8.ToArray();
        File.WriteAllBytes(archivePath, originalContent);
        var settings = new NativeQueryApprovalSettings(archivePath);

        var exception = Assert.Throws<IOException>(() =>
            settings.Write(required: false, isExcelRunning: () => false));

        Assert.Contains("original archive was not replaced", exception.Message, StringComparison.Ordinal);
        Assert.Contains("The backup is", exception.Message, StringComparison.Ordinal);
        Assert.Equal(originalContent, File.ReadAllBytes(archivePath));
        string[] files = Directory.GetFiles(folder.Path);
        Assert.Equal(2, files.Length);
        Assert.Contains(files, file => file.EndsWith(".bak", StringComparison.Ordinal));
    }

    [Fact]
    public void Set_RequiresConfirmationBeforeReadingOrChangingSetting()
    {
        var settings = new FakeApprovalSettings(new NativeQueryApprovalSnapshot(true, true));
        var commands = new SettingsCommands(null, settings, () => false);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            commands.SetNativeQueryApproval(required: false, confirm: false));

        Assert.Contains("confirm=true", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, settings.ReadCount);
        Assert.Equal(0, settings.WriteCount);
    }

    [Fact]
    public void Set_WhenApprovalRequirementIsOmitted_RefusesInsteadOfDisablingApproval()
    {
        var settings = new FakeApprovalSettings(new NativeQueryApprovalSnapshot(true, true));
        var commands = new SettingsCommands(null, settings, () => false);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            commands.SetNativeQueryApproval(required: null, confirm: true));

        Assert.Contains("required=true or required=false", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, settings.ReadCount);
        Assert.Equal(0, settings.WriteCount);
    }

    [Fact]
    public void Set_WhenExcelIsRunning_RefusesAndLeavesSettingUnchanged()
    {
        var settings = new FakeApprovalSettings(new NativeQueryApprovalSnapshot(true, true));
        var commands = new SettingsCommands(null, settings, () => true);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            commands.SetNativeQueryApproval(required: false, confirm: true));

        Assert.Contains("Close all Excel windows", exception.Message, StringComparison.Ordinal);
        Assert.True(settings.Snapshot.Required);
        Assert.Equal(0, settings.WriteCount);
    }

    [Fact]
    public void Set_WhenExcelStartsDuringArchivePreparation_LeavesOriginalArchiveUnchanged()
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        CreateArchive(archivePath, SettingsXml("l0"));
        byte[] originalArchive = File.ReadAllBytes(archivePath);
        int processCheckCount = 0;
        var settings = new NativeQueryApprovalSettings(archivePath);
        var commands = new SettingsCommands(
            null,
            settings,
            () => ++processCheckCount > 1);

        var exception = Assert.Throws<IOException>(() =>
            commands.SetNativeQueryApproval(required: false, confirm: true));

        Assert.Contains("Excel started while", exception.InnerException!.Message, StringComparison.Ordinal);
        Assert.Contains("The backup is", exception.Message, StringComparison.Ordinal);
        Assert.Equal(originalArchive, File.ReadAllBytes(archivePath));
        Assert.Equal(2, processCheckCount);
        Assert.Single(Directory.GetFiles(folder.Path, "*.bak"));
        Assert.Single(Directory.GetFiles(folder.Path, "User.zip"));
    }

    [Fact]
    public void Set_WhenValueAlreadyMatches_DoesNotCheckExcelOrCreateBackup()
    {
        var settings = new FakeApprovalSettings(new NativeQueryApprovalSnapshot(true, false));
        var commands = new SettingsCommands(null, settings, () => throw new InvalidOperationException("Must not run"));

        NativeQueryApprovalResult result = commands.SetNativeQueryApproval(required: true, confirm: true);

        Assert.True(result.Success);
        Assert.True(result.Required);
        Assert.False(result.ExplicitlySet);
        Assert.Equal(true, result.PreviousRequired);
        Assert.Null(result.BackupPath);
        Assert.Equal(0, settings.WriteCount);
    }

    [Theory]
    [InlineData("l0", false, "l1")]
    [InlineData("l1", true, "l0")]
    public void Set_CreatesBackupVerifiesNewValueAndPreservesOtherArchiveEntries(
        string oldValue,
        bool required,
        string newValue)
    {
        using var folder = new TemporaryFolder();
        string archivePath = Path.Combine(folder.Path, "User.zip");
        byte[] unrelatedContent = [0, 2, 4, 6, 8, 255];
        CreateArchive(archivePath, SettingsXml(oldValue), unrelatedContent);
        var settings = new NativeQueryApprovalSettings(archivePath);
        var commands = new SettingsCommands(null, settings, () => false);

        NativeQueryApprovalResult result = commands.SetNativeQueryApproval(required, confirm: true);

        Assert.True(result.Success);
        Assert.Equal(required, result.Required);
        Assert.True(result.ExplicitlySet);
        Assert.Equal(!required, result.PreviousRequired);
        Assert.Equal(archivePath, result.SettingsFilePath);
        Assert.Contains("all workbooks", result.Warning, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(result.BackupPath));
        foreach (string entryName in new[]
        {
            "[Content_Types].xml",
            "Credentials/Credentials.bin",
            "NativeQueries/NativeQueries.xml"
        })
        {
            byte[] expectedContent = entryName == "Credentials/Credentials.bin"
                ? unrelatedContent
                : ReadEntry(result.BackupPath!, entryName);
            Assert.Equal(expectedContent, ReadEntry(archivePath, entryName));
            Assert.Equal(expectedContent, ReadEntry(result.BackupPath!, entryName));
        }
        Assert.Equal(
            SettingsXml(oldValue),
            ReadEntryText(result.BackupPath!, "UserInterface/Settings.xml"));
        Assert.Equal(
            newValue,
            (string?)ReadEntryXml(archivePath, "UserInterface/Settings.xml")
                .Descendants("Entry")
                .Single(entry => (string?)entry.Attribute("Type") == "DisableNativeDbQueries")
                .Attribute("Value"));
    }

    [Fact]
    public void Set_WhenReadbackDoesNotMatch_ReportsThatArchiveWasNotVerified()
    {
        var settings = new FakeApprovalSettings(new NativeQueryApprovalSnapshot(true, true))
        {
            PersistOnWrite = false
        };
        var commands = new SettingsCommands(null, settings, () => false);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            commands.SetNativeQueryApproval(required: false, confirm: true));

        Assert.Contains("verification returned required=True", exception.Message, StringComparison.Ordinal);
        Assert.Contains("backup.zip", exception.Message, StringComparison.Ordinal);
        Assert.Equal(1, settings.WriteCount);
    }

    private static string SettingsXml(string? value)
    {
        string setting = value is null
            ? """<Entry Type="KeepThis" Value="unchanged" />"""
            : $"""<Entry Type="DisableNativeDbQueries" Value="{value}" /><Entry Type="KeepThis" Value="unchanged" />""";
        return $"<UISettingsConfig><Entries>{setting}</Entries></UISettingsConfig>";
    }

    private static void CreateArchive(string path, string? settingsXml, byte[]? unrelatedContent = null)
    {
        using var file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        WriteEntry(archive, "[Content_Types].xml", "<Types />"u8.ToArray());
        WriteEntry(archive, "Credentials/Credentials.bin", unrelatedContent ?? [10, 20, 30]);
        WriteEntry(archive, "NativeQueries/NativeQueries.xml", "<NativeQueries />"u8.ToArray());
        if (settingsXml is not null)
        {
            WriteEntry(archive, "UserInterface/Settings.xml", Encoding.UTF8.GetBytes(settingsXml));
        }
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] content)
    {
        using Stream destination = archive.CreateEntry(name).Open();
        destination.Write(content);
    }

    private static byte[] ReadEntry(string archivePath, string entryName)
    {
        using var file = File.OpenRead(archivePath);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        using Stream source = archive.GetEntry(entryName)!.Open();
        using var memory = new MemoryStream();
        source.CopyTo(memory);
        return memory.ToArray();
    }

    private static string ReadEntryText(string archivePath, string entryName) =>
        Encoding.UTF8.GetString(ReadEntry(archivePath, entryName));

    private static XDocument ReadEntryXml(string archivePath, string entryName)
    {
        using var file = File.OpenRead(archivePath);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        using Stream source = archive.GetEntry(entryName)!.Open();
        return XDocument.Load(source);
    }

    private sealed class FakeApprovalSettings(NativeQueryApprovalSnapshot snapshot) : INativeQueryApprovalSettings
    {
        public NativeQueryApprovalSnapshot Snapshot { get; private set; } = snapshot;
        public string SettingsFilePath => "C:\\Users\\test\\User.zip";
        public int ReadCount { get; private set; }
        public int WriteCount { get; private set; }
        public bool PersistOnWrite { get; init; } = true;

        public NativeQueryApprovalSnapshot Read()
        {
            ReadCount++;
            return Snapshot;
        }

        public string Write(bool required, Func<bool> isExcelRunning)
        {
            WriteCount++;
            if (PersistOnWrite)
            {
                Snapshot = new NativeQueryApprovalSnapshot(required, true);
            }
            return "backup.zip";
        }
    }

    private sealed class TemporaryFolder : IDisposable
    {
        public TemporaryFolder()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"native-query-approval-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
