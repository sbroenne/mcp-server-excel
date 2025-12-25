using System.Reflection;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Loads markdown content from embedded resources for MCP prompts and elicitations.
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
}
