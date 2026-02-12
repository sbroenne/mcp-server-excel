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

    /// <summary>Path to the Scriban template file (.sbn)</summary>
    [Required]
    public string TemplatePath { get; set; } = "";

    /// <summary>Output path for the generated file</summary>
    [Required]
    public string OutputPath { get; set; } = "";

    /// <summary>Version number to inject into template</summary>
    [Required]
    public string Version { get; set; } = "";

    /// <summary>Path to the generated _SkillManifest.g.cs file containing JSON metadata</summary>
    public string? ManifestPath { get; set; }

    /// <summary>Repository URL to inject into template</summary>
    public string? Repository { get; set; }

    /// <summary>Documentation URL to inject into template</summary>
    public string? Documentation { get; set; }

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
        var model = new SkillTemplateModel
        {
            Version = Version,
            Repository = Repository ?? "https://github.com/sbroenne/mcp-server-excel",
            Documentation = Documentation ?? "https://excelmcpserver.dev/"
        };

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
                model.ToolCount = manifest.TotalCommands;
                model.OperationCount = manifest.TotalOperations;
                model.CliCommands = manifest.Commands?.Select(c => new CliCommand
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

                Log.LogMessage(MessageImportance.Normal, $"Loaded manifest: {model.ToolCount} commands, {model.OperationCount} operations");
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
