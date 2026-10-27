using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Workers;

internal sealed class WorkerRegistry
{
    internal static string Build => typeof(WorkerRegistry).Assembly.ManifestModule.ModuleVersionId.ToString("N");
    private readonly string _directory;
    internal Action<string>? BeforeRecordReadForTests { get; set; }

    internal WorkerRegistry(string? directory = null)
    {
        _directory = directory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelMcp", "workers");
        var security = new DirectorySecurity();
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        security.AddAccessRule(new FileSystemAccessRule(
            WindowsIdentity.GetCurrent().User!,
            FileSystemRights.FullControl,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
            PropagationFlags.None,
            AccessControlType.Allow));
        FileSystemAclExtensions.CreateDirectory(security, _directory);
    }

    internal void Write(WorkerRegistration registration)
    {
        var path = GetPath(registration.PipeName);
        var temporary = path + ".tmp";
        File.WriteAllText(temporary, JsonSerializer.Serialize(registration, ServiceProtocol.JsonOptions));
        File.Move(temporary, path, overwrite: true);
    }

    internal IReadOnlyList<WorkerRegistration> Read()
    {
        var registrations = new List<WorkerRegistration>();
        foreach (var path in Directory.EnumerateFiles(_directory, "*.json"))
        {
            BeforeRecordReadForTests?.Invoke(path);
            WorkerRegistration registration;
            try
            {
                if (new FileInfo(path).Length > 65536)
                    throw new InvalidDataException("An Excel worker record exceeds its size limit.");
                registration = JsonSerializer.Deserialize<WorkerRegistration>(
                    File.ReadAllText(path), ServiceProtocol.JsonOptions)
                    ?? throw new InvalidDataException("An Excel worker record is empty.");
            }
            catch (FileNotFoundException)
            {
                // A worker can remove its record after directory enumeration.
                continue;
            }
            _ = GetPath(registration.PipeName);
            if (registration.Worker.IsAlive()
                || registration.Excel is { } excel && !OwnedProcessGuard.TryConfirmExited(excel))
            {
                if (registrations.Count == 128)
                    throw new InvalidOperationException("Too many retained Excel workers. Inspect the retained processes before retrying.");
                registrations.Add(registration);
            }
            else
                File.Delete(path);
        }
        return registrations;
    }

    internal void Remove(string pipeName) => File.Delete(GetPath(pipeName));

    private string GetPath(string pipeName)
    {
        if (!pipeName.StartsWith("excelmcp-worker-", StringComparison.Ordinal)
            || !Guid.TryParseExact(pipeName["excelmcp-worker-".Length..], "N", out _))
            throw new InvalidDataException("Invalid Excel worker record name.");
        return Path.Combine(_directory, pipeName + ".json");
    }
}
