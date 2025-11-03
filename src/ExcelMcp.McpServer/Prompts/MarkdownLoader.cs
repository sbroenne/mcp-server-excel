using System.Reflection;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Loads markdown content from embedded resources for MCP prompts, completions, and elicitations.
/// Follows guidance from mcp-llm-guidance.instructions.md for markdown-based LLM guidance.
/// </summary>
public static class MarkdownLoader
{
    private static readonly Assembly _assembly = typeof(MarkdownLoader).Assembly;
    private static readonly string _baseNamespace = "Sbroenne.ExcelMcp.McpServer.Prompts.Content";

    /// <summary>
    /// Load a prompt markdown file from Content/ directory
    /// </summary>
    /// <param name="fileName">File name without path (e.g., "excel_powerquery.md")</param>
    /// <returns>Markdown content</returns>
    public static string LoadPrompt(string fileName)
    {
        return LoadMarkdownFile($"{_baseNamespace}.{fileName}");
    }

    /// <summary>
    /// Load a completion markdown file from Content/Completions/ directory
    /// </summary>
    /// <param name="fileName">File name without path (e.g., "action_powerquery.md")</param>
    /// <returns>Markdown content as newline-separated values</returns>
    public static string LoadCompletion(string fileName)
    {
        return LoadMarkdownFile($"{_baseNamespace}.Completions.{fileName}");
    }

    /// <summary>
    /// Load a completion as list of values (one per line)
    /// </summary>
    /// <param name="fileName">File name without path</param>
    /// <returns>List of completion values</returns>
    public static List<string> LoadCompletionValues(string fileName)
    {
        var content = LoadCompletion(fileName);
        return content
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line) && !line.StartsWith("#"))
            .ToList();
    }

    /// <summary>
    /// Load an elicitation markdown file from Content/Elicitations/ directory
    /// </summary>
    /// <param name="fileName">File name without path (e.g., "powerquery_import.md")</param>
    /// <returns>Markdown content</returns>
    public static string LoadElicitation(string fileName)
    {
        return LoadMarkdownFile($"{_baseNamespace}.Elicitations.{fileName}");
    }

    /// <summary>
    /// Load markdown file from embedded resource
    /// </summary>
    private static string LoadMarkdownFile(string resourceName)
    {
        try
        {
            using var stream = _assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                throw new FileNotFoundException(
                    $"Embedded resource not found: {resourceName}. " +
                    $"Ensure the .md file is marked as EmbeddedResource in .csproj");
            }

            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to load markdown file: {resourceName}. Error: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// List all available embedded markdown resources (for debugging)
    /// </summary>
    public static List<string> ListEmbeddedResources()
    {
        return _assembly.GetManifestResourceNames()
            .Where(name => name.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
            .ToList();
    }
}
