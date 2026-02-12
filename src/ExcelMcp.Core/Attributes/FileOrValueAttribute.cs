namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Indicates that this parameter can be provided either as a direct value or via a file path.
/// The generator will create two parameters in MCP/CLI: the base param and a file param.
/// Example: [FileOrValue] string mCode → generates mCode + mCodeFile parameters
/// </summary>
[AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
public sealed class FileOrValueAttribute : Attribute
{
    /// <summary>
    /// The suffix to use for the file parameter. Default is "File".
    /// Example: mCode with suffix "File" → mCodeFile
    /// </summary>
    public string FileSuffix { get; }

    /// <summary>
    /// Creates a FileOrValue attribute with default "File" suffix.
    /// </summary>
    public FileOrValueAttribute() : this("File") { }

    /// <summary>
    /// Creates a FileOrValue attribute with a custom suffix.
    /// </summary>
    /// <param name="fileSuffix">Suffix for the file parameter name</param>
    public FileOrValueAttribute(string fileSuffix)
    {
        FileSuffix = fileSuffix ?? throw new ArgumentNullException(nameof(fileSuffix));
    }
}
