using Microsoft.Build.Framework;
using Scriban;
using Scriban.Runtime;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Build.Tasks;

/// <summary>
/// MSBuild task that generates skill files from Scriban templates.
/// </summary>
public class GenerateSkillFile : Microsoft.Build.Utilities.Task
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    private static readonly char[] CommandSeparators = { ';', ',' };

    /// <summary>Path to the Scriban template file (.sbn)</summary>
    [Required]
    public string TemplatePath { get; set; } = "";

    /// <summary>Output path for the generated file</summary>
    [Required]
    public string OutputPath { get; set; } = "";

    /// <summary>Path to the generated _SkillManifest.g.cs file containing JSON metadata</summary>
    public string? ManifestPath { get; set; }

    /// <summary>
    /// Semicolon-separated command names to EXCLUDE from the skill surface and counts
    /// (e.g. "diag" — a CLI-only self-test that is not part of the user-facing surface).
    /// </summary>
    public string? ExcludeCommands { get; set; }

    /// <summary>
    /// Extra operations to ADD to the operation count for hand-written tools that are
    /// not Core [ServiceCategory] classes and therefore absent from the manifest
    /// (e.g. the file/session tool backed by FileAction). Default 0.
    /// </summary>
    public int ExtraOperationCount { get; set; }

    /// <summary>
    /// Extra tools to ADD to the tool count for hand-written tools absent from the
    /// manifest (e.g. the file/session tool). Default 0.
    /// </summary>
    public int ExtraToolCount { get; set; }

    /// <summary>Executes the task to generate the skill file from the template.</summary>
    /// <returns>true if the task succeeded; otherwise, false.</returns>
    public override bool Execute()
    {
        try
        {
            if (!File.Exists(TemplatePath))
            {
                Log.LogError($"Template file not found: {TemplatePath}");
                return false;
            }

            // Read and parse template
            var templateContent = File.ReadAllText(TemplatePath);
            var template = Template.Parse(templateContent);

            if (template.HasErrors)
            {
                foreach (var error in template.Messages)
                {
                    Log.LogError($"Template error: {error}");
                }
                return false;
            }

            // Build model from manifest
            var model = BuildModelFromManifest();

            // Render template
            var scriptObject = new ScriptObject();
            scriptObject.Import(model, renamer: member => member.Name.ToLowerInvariant());

            var context = new TemplateContext();
            context.PushGlobal(scriptObject);

            var output = template.Render(context);

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(OutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write output
            File.WriteAllText(OutputPath, output);
            Log.LogMessage(MessageImportance.High, $"Generated: {OutputPath}");

            return true;
        }
        catch (Exception ex)
        {
            Log.LogErrorFromException(ex, showStackTrace: true);
            return false;
        }
    }

    private SkillTemplateModel BuildModelFromManifest()
    {
        var model = new SkillTemplateModel();

        if (string.IsNullOrEmpty(ManifestPath) || !File.Exists(ManifestPath))
        {
            Log.LogWarning($"Manifest file not found: {ManifestPath}. Skill will have no command reference.");
            return model;
        }

        // Read the generated _SkillManifest.g.cs file and extract JSON
        var manifestContent = File.ReadAllText(ManifestPath);
        var json = ExtractJsonFromManifest(manifestContent);

        if (string.IsNullOrEmpty(json))
        {
            Log.LogWarning($"Could not extract JSON from manifest: {ManifestPath}");
            return model;
        }

        // Parse JSON
        try
        {
            var manifest = JsonSerializer.Deserialize<SkillManifest>(json!, JsonOptions);
            if (manifest != null)
            {
                var commands = manifest.Commands ?? new List<ManifestCommand>();

                // Exclude CLI-only / non-user-facing commands (e.g. "diag") from the
                // surface and counts so the skill reflects the real user-facing tools.
                var excluded = (ExcludeCommands ?? string.Empty)
                    .Split(CommandSeparators, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToList();
                bool IsExcluded(string? name) =>
                    name != null && excluded.Any(e => string.Equals(e, name, StringComparison.OrdinalIgnoreCase));

                var excludedCommands = commands.Where(c => IsExcluded(c.Name)).ToList();
                var includedCommands = commands.Where(c => !IsExcluded(c.Name)).ToList();

                var excludedOps = excludedCommands.Sum(c => c.Actions?.Length ?? 0);

                // Tool count = manifest total - excluded + hand-written extras (e.g. file tool).
                model.ToolCount = manifest.TotalCommands - excludedCommands.Count + ExtraToolCount;
                // Operation count = manifest total - excluded ops + hand-written extras (e.g. file ops).
                model.OperationCount = manifest.TotalOperations - excludedOps + ExtraOperationCount;

                model.CliCommands = includedCommands.Select(c => new CliCommand
                {
                    Name = c.Name ?? "",
                    Description = c.Description ?? "",
                    Actions = c.Actions?.ToList() ?? new List<string>(),
                    Parameters = c.Parameters?.Select(p => new CliParameter
                    {
                        Name = p.Name ?? "",
                        Description = p.Description ?? ""
                    }).ToList() ?? new List<CliParameter>()
                }).ToList();

                Log.LogMessage(MessageImportance.Normal, $"Loaded manifest: {model.ToolCount} commands, {model.OperationCount} operations (excluded: {excludedCommands.Count} cmds/{excludedOps} ops, extra: {ExtraToolCount} tools/{ExtraOperationCount} ops)");
            }
        }
        catch (JsonException ex)
        {
            Log.LogWarning($"Failed to parse manifest JSON: {ex.Message}");
        }

        return model;
    }

    private static string? ExtractJsonFromManifest(string content)
    {
        // The manifest file contains: public const string Json = @"{...}";
        // We need to extract the JSON between @" and ";
        const string startMarker = "public const string Json = @\"";
        const string endMarker = "\";";

        var startIndex = content.IndexOf(startMarker, StringComparison.Ordinal);
        if (startIndex < 0)
            return null;

        startIndex += startMarker.Length;

        var endIndex = content.LastIndexOf(endMarker, StringComparison.Ordinal);
        if (endIndex <= startIndex)
            return null;

        var json = content.Substring(startIndex, endIndex - startIndex);

        // The JSON uses doubled quotes ("") for escaping in verbatim string
        // Convert back to regular JSON quotes
        json = json.Replace("\"\"", "\"");

        return json;
    }
}

/// <summary>JSON manifest structure from the generator.</summary>
internal sealed class SkillManifest
{
    public List<ManifestCommand>? Commands { get; set; }
    public int TotalCommands { get; set; }
    public int TotalOperations { get; set; }
}

/// <summary>Command entry in the manifest.</summary>
internal sealed class ManifestCommand
{
    public string? Name { get; set; }
    public string? McpTool { get; set; }
    public string? Description { get; set; }
    public string[]? Actions { get; set; }
    public ManifestParameter[]? Parameters { get; set; }
}

/// <summary>Parameter entry in the manifest.</summary>
internal sealed class ManifestParameter
{
    public string? Name { get; set; }
    public string? Description { get; set; }
}
