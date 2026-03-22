using System.Text.Json;

namespace Sbroenne.ExcelMcp.CLI.Tests.Helpers;

internal sealed class CliConfigUpdateDefinition
{
    public string SheetName { get; init; } = string.Empty;

    public string RangeAddress { get; init; } = string.Empty;

    public string ValuesJson { get; init; } = string.Empty;

    public int TimeoutMs { get; init; } = 30000;
}

internal sealed class CliPowerQueryRefreshDefinition
{
    public string QueryName { get; init; } = string.Empty;

    public bool? ExpectedSuccess { get; init; }

    public int TimeoutMs { get; init; } = 120000;
}

internal sealed class CliPowerQueryWorkflowDefinition
{
    public string SourceWorkbookPath { get; init; } = string.Empty;

    public string WorkingCopyFileStem { get; init; } = "anonymized-refresh-copy";

    public CliConfigUpdateDefinition ConfigUpdate { get; init; } = new();

    public IReadOnlyList<CliPowerQueryRefreshDefinition> RefreshSequence { get; init; } = [];
}

internal sealed class CliPowerQueryWorkflowFixture : IDisposable
{
    private const string LocalDefinitionFileName = "anonymized-serial-refresh-scenario.local.json";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };
    private readonly string _tempDirectory;

    public CliPowerQueryWorkflowFixture()
    {
        Definition = LoadDefinition();
        _tempDirectory = Path.Combine(Path.GetTempPath(), $"CliWorkflow_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDirectory);

        var workbookExtension = Path.GetExtension(Definition.SourceWorkbookPath);
        WorkingCopyPath = Path.Combine(
            _tempDirectory,
            $"{Definition.WorkingCopyFileStem}-{Guid.NewGuid():N}{workbookExtension}");

        ResetWorkingCopy();
    }

    public CliPowerQueryWorkflowDefinition Definition { get; }

    public string WorkingCopyPath { get; private set; }

    public void ResetWorkingCopy()
    {
        if (File.Exists(WorkingCopyPath))
        {
            File.Delete(WorkingCopyPath);
        }

        File.Copy(Definition.SourceWorkbookPath, WorkingCopyPath);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDirectory))
        {
            try
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }

    public static JsonElement ExtractFirstScalarValue(string valuesJson)
    {
        using var document = JsonDocument.Parse(valuesJson);
        return document.RootElement[0][0].Clone();
    }

    private static CliPowerQueryWorkflowDefinition LoadDefinition()
    {
        string repoRoot = GetRepoRoot();
        string definitionPath = Path.Combine(repoRoot, "TestResults", "real-workbook-repro", LocalDefinitionFileName);

        if (!File.Exists(definitionPath))
        {
            throw new FileNotFoundException(
                $"Local workflow definition not found: {definitionPath}{Environment.NewLine}" +
                "Copy tests\\ExcelMcp.CLI.Tests\\Integration\\TestAssets\\anonymized-serial-refresh-scenario.sample.json " +
                "to TestResults\\real-workbook-repro\\anonymized-serial-refresh-scenario.local.json and fill in local values.");
        }

        var definition = JsonSerializer.Deserialize<CliPowerQueryWorkflowDefinition>(
            File.ReadAllText(definitionPath),
            JsonOptions);

        if (definition == null ||
            string.IsNullOrWhiteSpace(definition.SourceWorkbookPath) ||
            string.IsNullOrWhiteSpace(definition.WorkingCopyFileStem) ||
            string.IsNullOrWhiteSpace(definition.ConfigUpdate.SheetName) ||
            string.IsNullOrWhiteSpace(definition.ConfigUpdate.RangeAddress) ||
            string.IsNullOrWhiteSpace(definition.ConfigUpdate.ValuesJson) ||
            definition.RefreshSequence.Count == 0)
        {
            throw new InvalidOperationException(
                "Local workflow definition is invalid. Provide sourceWorkbookPath, workingCopyFileStem, configUpdate, and refreshSequence.");
        }

        string resolvedSourcePath = definition.SourceWorkbookPath;
        if (!Path.IsPathRooted(resolvedSourcePath))
        {
            resolvedSourcePath = Path.GetFullPath(Path.Combine(repoRoot, resolvedSourcePath));
        }

        if (!File.Exists(resolvedSourcePath))
        {
            throw new FileNotFoundException(
                $"Local source workbook not found: {resolvedSourcePath}{Environment.NewLine}" +
                "Update sourceWorkbookPath in the local workflow definition.");
        }

        return new CliPowerQueryWorkflowDefinition
        {
            SourceWorkbookPath = resolvedSourcePath,
            WorkingCopyFileStem = definition.WorkingCopyFileStem,
            ConfigUpdate = definition.ConfigUpdate,
            RefreshSequence = definition.RefreshSequence
        };
    }

    private static string GetRepoRoot()
    {
        var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            dir = dir.Parent;
        }

        return dir?.FullName ?? throw new InvalidOperationException("Could not find repo root.");
    }
}
