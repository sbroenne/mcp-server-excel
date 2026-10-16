using System.Diagnostics;
using System.IO.Compression;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Core.Commands.Settings;

internal sealed record NativeQueryApprovalSnapshot(bool Required, bool ExplicitlySet);

internal interface INativeQueryApprovalSettings
{
    string SettingsFilePath { get; }

    NativeQueryApprovalSnapshot Read();

    string Write(bool required, Func<bool> isExcelRunning);
}

/// <summary>
/// Reads and updates the current user's Power Query settings archive. The archive is not an Excel workbook.
/// </summary>
internal sealed class NativeQueryApprovalSettings : INativeQueryApprovalSettings
{
    private const string SettingsEntryName = "UserInterface/Settings.xml";
    private const string SettingName = "DisableNativeDbQueries";
    private readonly string _settingsFilePath;

    internal NativeQueryApprovalSettings()
        : this(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "Office", "16.0", "PowerQuery", "User.zip"))
    {
    }

    internal NativeQueryApprovalSettings(string settingsFilePath) =>
        _settingsFilePath = settingsFilePath;

    public string SettingsFilePath => _settingsFilePath;

    public NativeQueryApprovalSnapshot Read()
    {
        EnsureWindows();

        using var stream = OpenSettingsArchive();
        ZipArchive archive;
        try
        {
            archive = new ZipArchive(stream, ZipArchiveMode.Read);
        }
        catch (InvalidDataException exception)
        {
            throw new InvalidDataException(
                $"The Power Query settings archive at '{_settingsFilePath}' is not a readable Office 16.0 archive.",
                exception);
        }

        using (archive)
        {
            var xml = ReadSettingsXml(archive);
            var entry = FindSetting(xml);
            if (entry is null)
            {
                return new NativeQueryApprovalSnapshot(Required: true, ExplicitlySet: false);
            }

            string? value = (string?)entry.Attribute("Value");
            return value switch
            {
                "l0" => new NativeQueryApprovalSnapshot(Required: true, ExplicitlySet: true),
                "l1" => new NativeQueryApprovalSnapshot(Required: false, ExplicitlySet: true),
                _ => throw new InvalidDataException(
                    $"The {SettingName} value in '{_settingsFilePath}' is unsupported. Expected 'l0' or 'l1'.")
            };
        }
    }

    public string Write(bool required, Func<bool> isExcelRunning)
    {
        EnsureWindows();

        string backupPath = $"{_settingsFilePath}.{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}.bak";
        string backupTemporaryPath = $"{backupPath}.{Guid.NewGuid():N}.tmp";
        string temporaryPath = $"{_settingsFilePath}.{Guid.NewGuid():N}.tmp";
        bool backupCreated = false;
        using FileStream source = OpenSettingsArchive();
        try
        {
            using (var backup = new FileStream(
                backupTemporaryPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None))
            {
                source.CopyTo(backup);
                backup.Flush(flushToDisk: true);
            }
            File.Move(backupTemporaryPath, backupPath, overwrite: false);
            backupCreated = true;

            source.Position = 0;
            using (var temporary = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var sourceArchive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
            {
                using (var destinationArchive = new ZipArchive(temporary, ZipArchiveMode.Create, leaveOpen: true))
                {
                    CopyArchive(sourceArchive, destinationArchive, required);
                }
                temporary.Flush(flushToDisk: true);
            }
            source.Dispose();
            if (isExcelRunning())
            {
                throw new IOException(
                    "Excel started while the native-query approval archive was being prepared. The original archive was not replaced.");
            }

            File.Move(temporaryPath, _settingsFilePath, overwrite: true);
            return backupPath;
        }
        catch (InvalidDataException exception)
        {
            throw CreateWriteFailure(exception, temporaryPath, backupTemporaryPath, backupPath, backupCreated);
        }
        catch (System.Xml.XmlException exception)
        {
            throw CreateWriteFailure(exception, temporaryPath, backupTemporaryPath, backupPath, backupCreated);
        }
        catch (IOException exception)
        {
            throw CreateWriteFailure(exception, temporaryPath, backupTemporaryPath, backupPath, backupCreated);
        }
        catch (UnauthorizedAccessException exception)
        {
            throw CreateWriteFailure(exception, temporaryPath, backupTemporaryPath, backupPath, backupCreated);
        }
        catch (ArgumentException exception)
        {
            throw CreateWriteFailure(exception, temporaryPath, backupTemporaryPath, backupPath, backupCreated);
        }
        catch (NotSupportedException exception)
        {
            throw CreateWriteFailure(exception, temporaryPath, backupTemporaryPath, backupPath, backupCreated);
        }
    }

    internal static bool IsExcelRunning()
    {
        Process[] processes = Process.GetProcesses();
        try
        {
            return processes.Any(process =>
                string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            foreach (Process process in processes)
            {
                process.Dispose();
            }
        }
    }

    private static void EnsureWindows()
    {
        if (!OperatingSystem.IsWindows())
        {
            throw new NotSupportedException("Native query approval settings are supported only on Windows Office 16.0.");
        }
    }

    private FileStream OpenSettingsArchive()
    {
        try
        {
            return File.OpenRead(_settingsFilePath);
        }
        catch (FileNotFoundException exception)
        {
            throw new FileNotFoundException(
                $"The Power Query settings archive for supported Office 16.0 installations was not found at '{_settingsFilePath}'.",
                _settingsFilePath,
                exception);
        }
        catch (DirectoryNotFoundException exception)
        {
            throw new FileNotFoundException(
                $"The Power Query settings archive for supported Office 16.0 installations was not found at '{_settingsFilePath}'.",
                _settingsFilePath,
                exception);
        }
    }

    private IOException CreateWriteFailure(
        Exception exception,
        string temporaryPath,
        string backupTemporaryPath,
        string backupPath,
        bool backupCreated)
    {
        Exception reportedException = exception;
        reportedException = TryDeleteTemporaryFile(temporaryPath, reportedException);
        reportedException = TryDeleteTemporaryFile(backupTemporaryPath, reportedException);

        string backupMessage = backupCreated
            ? $" The backup is '{backupPath}'."
            : " No complete backup was created.";
        string temporaryMessage = reportedException == exception
            ? string.Empty
            : $" Temporary files may remain at '{temporaryPath}' or '{backupTemporaryPath}'.";
        return new IOException(
            $"Could not update '{_settingsFilePath}'. The original archive was not replaced.{backupMessage}{temporaryMessage}",
            reportedException);
    }

    private static Exception TryDeleteTemporaryFile(string path, Exception reportedException)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException cleanupException)
        {
            return new AggregateException(reportedException, cleanupException);
        }
        catch (UnauthorizedAccessException cleanupException)
        {
            return new AggregateException(reportedException, cleanupException);
        }

        return reportedException;
    }

    private static XDocument ReadSettingsXml(ZipArchive archive)
    {
        ZipArchiveEntry[] entries = archive.Entries
            .Where(entry => string.Equals(entry.FullName, SettingsEntryName, StringComparison.Ordinal))
            .ToArray();
        if (entries.Length != 1)
        {
            throw new InvalidDataException(
                $"Expected exactly one '{SettingsEntryName}' entry in the Power Query settings archive; found {entries.Length}.");
        }

        using Stream stream = entries[0].Open();
        XDocument document;
        try
        {
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is System.Xml.XmlException or IOException or InvalidDataException)
        {
            throw new InvalidDataException($"The '{SettingsEntryName}' settings entry could not be read.", exception);
        }

        ValidateStructure(document);

        return document;
    }

    private static XElement? FindSetting(XDocument document)
    {
        XElement entries = document.Root!.Element("Entries")!;
        XElement[] matches = entries.Elements("Entry")
            .Where(entry => string.Equals((string?)entry.Attribute("Type"), SettingName, StringComparison.Ordinal))
            .ToArray();
        if (matches.Length > 1)
        {
            throw new InvalidDataException($"The {SettingName} setting appears more than once.");
        }

        return matches.SingleOrDefault();
    }

    private static void CopyArchive(ZipArchive source, ZipArchive destination, bool required)
    {
        ZipArchiveEntry[] settingsEntries = source.Entries
            .Where(entry => string.Equals(entry.FullName, SettingsEntryName, StringComparison.Ordinal))
            .ToArray();
        if (settingsEntries.Length != 1)
        {
            throw new InvalidDataException(
                $"Expected exactly one '{SettingsEntryName}' entry in the Power Query settings archive; found {settingsEntries.Length}.");
        }

        foreach (ZipArchiveEntry sourceEntry in source.Entries)
        {
            ZipArchiveEntry destinationEntry = destination.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
            destinationEntry.LastWriteTime = sourceEntry.LastWriteTime;
            destinationEntry.ExternalAttributes = sourceEntry.ExternalAttributes;
            using Stream sourceStream = sourceEntry.Open();
            using Stream destinationStream = destinationEntry.Open();
            if (string.Equals(sourceEntry.FullName, SettingsEntryName, StringComparison.Ordinal))
            {
                XDocument document;
                try
                {
                    document = XDocument.Load(sourceStream, LoadOptions.PreserveWhitespace);
                }
                catch (Exception exception) when (exception is System.Xml.XmlException or IOException or InvalidDataException)
                {
                    throw new InvalidDataException($"The '{SettingsEntryName}' settings entry could not be read.", exception);
                }

                ValidateStructure(document);
                XElement? entry = FindSetting(document);
                if (entry is null)
                {
                    document.Root!.Element("Entries")!.Add(new XElement(
                        "Entry",
                        new XAttribute("Type", SettingName),
                        new XAttribute("Value", required ? "l0" : "l1")));
                }
                else
                {
                    string? value = (string?)entry.Attribute("Value");
                    if (value is not ("l0" or "l1"))
                    {
                        throw new InvalidDataException(
                            $"The {SettingName} value in the settings archive is unsupported. Expected 'l0' or 'l1'.");
                    }
                    entry.SetAttributeValue("Value", required ? "l0" : "l1");
                }

                document.Save(destinationStream, SaveOptions.DisableFormatting);
            }
            else
            {
                sourceStream.CopyTo(destinationStream);
            }
        }
    }

    private static void ValidateStructure(XDocument document)
    {
        if (document.Root?.Name != XName.Get("UISettingsConfig") ||
            document.Root.Elements("Entries").Count() != 1)
        {
            throw new InvalidDataException(
                $"The '{SettingsEntryName}' entry has an unsupported settings structure.");
        }
    }
}
