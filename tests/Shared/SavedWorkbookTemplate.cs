using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Tests.Infrastructure;

internal sealed class SavedWorkbookTemplateStore
{
    private readonly string _templateDirectory;
    private readonly Action<string> _createTemplate;
    private readonly ConcurrentDictionary<TemplateKey, Lazy<string>> _templates = new();

    internal SavedWorkbookTemplateStore(string templateDirectory, Action<string> createTemplate)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templateDirectory);
        ArgumentNullException.ThrowIfNull(createTemplate);
        _templateDirectory = templateDirectory;
        _createTemplate = createTemplate;
    }

    internal string CopyTo(string destinationPath, string baselineName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(baselineName);

        var extension = Path.GetExtension(destinationPath);
        if (string.IsNullOrWhiteSpace(extension))
        {
            throw new ArgumentException("A workbook extension is required.", nameof(destinationPath));
        }

        var key = new TemplateKey(baselineName, extension.ToLowerInvariant());
        var templatePath = _templates.GetOrAdd(
            key,
            static (templateKey, state) => new Lazy<string>(
                () => state.CreateTemplate(templateKey),
                LazyThreadSafetyMode.ExecutionAndPublication),
            this).Value;

        var destinationDirectory = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrEmpty(destinationDirectory))
        {
            Directory.CreateDirectory(destinationDirectory);
        }

        File.Copy(templatePath, destinationPath, overwrite: false);
        return destinationPath;
    }

    private string CreateTemplate(TemplateKey key)
    {
        Directory.CreateDirectory(_templateDirectory);
        var safeName = string.Concat(key.BaselineName.Select(character =>
            Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        var templatePath = Path.Combine(_templateDirectory, $"{safeName}_{Guid.NewGuid():N}{key.Extension}");
        try
        {
            _createTemplate(templatePath);
            if (!File.Exists(templatePath))
            {
                throw new InvalidOperationException(
                    $"The template creator did not create '{templatePath}'.");
            }

            return templatePath;
        }
        catch
        {
            File.Delete(templatePath);
            throw;
        }
    }

    private readonly record struct TemplateKey(string BaselineName, string Extension);
}

internal static class SavedWorkbookTemplates
{
    private static readonly string TemplateDirectory = Path.Combine(
        Path.GetTempPath(),
        $"ExcelMcpSavedWorkbookTemplates_{Environment.ProcessId}_{Guid.NewGuid():N}");

    private static readonly SavedWorkbookTemplateStore BlankTemplates =
        new(TemplateDirectory, CreateBlankWorkbook);

    static SavedWorkbookTemplates()
    {
        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try
            {
                if (Directory.Exists(TemplateDirectory))
                {
                    Directory.Delete(TemplateDirectory, recursive: true);
                }
            }
            catch
            {
                // Process exit cleanup is best-effort.
            }
        };
    }

    internal static string CopyBlankTo(string destinationPath) =>
        BlankTemplates.CopyTo(destinationPath, "blank");

    internal static SavedWorkbookTemplateStore CreateStore(Action<string> createTemplate) =>
        new(TemplateDirectory, createTemplate);

    private static void CreateBlankWorkbook(string templatePath)
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(templatePath, show: false);
        manager.CloseSession(sessionId, save: true);
    }
}
