namespace Sbroenne.ExcelMcp.Generators.Common;

/// <summary>
/// Extracted service information from an interface marked with [ServiceCategory].
/// </summary>
public sealed class ServiceInfo
{
    public string Category { get; }
    public string CategoryPascal { get; }
    public string McpToolName { get; }
    public bool NoSession { get; }
    public List<MethodInfo> Methods { get; }

    public ServiceInfo(string category, string categoryPascal, string mcpToolName, bool noSession, List<MethodInfo> methods)
    {
        Category = category;
        CategoryPascal = categoryPascal;
        McpToolName = mcpToolName;
        NoSession = noSession;
        Methods = methods;
    }
}

/// <summary>
/// Extracted method information from interface method.
/// </summary>
public sealed class MethodInfo
{
    public string MethodName { get; }
    public string ActionName { get; }
    public string ReturnType { get; }
    public string McpTool { get; }
    public List<ParameterInfo> Parameters { get; }
    public string? XmlDocSummary { get; }

    public MethodInfo(string methodName, string actionName, string returnType, string mcpTool,
        List<ParameterInfo> parameters, string? xmlDocSummary = null)
    {
        MethodName = methodName;
        ActionName = actionName;
        ReturnType = returnType;
        McpTool = mcpTool;
        Parameters = parameters;
        XmlDocSummary = xmlDocSummary;
    }
}

/// <summary>
/// Extracted parameter information.
/// </summary>
public sealed class ParameterInfo
{
    public string Name { get; }
    public string TypeName { get; }
    public bool HasDefault { get; }
    public string? DefaultValue { get; }
    public bool IsFileOrValue { get; }
    public string? FileSuffix { get; }
    public bool IsFromString { get; }
    public string? ExposedName { get; }
    public bool IsRequired { get; }
    public bool IsEnum { get; }
    public string? XmlDocDescription { get; }

    public ParameterInfo(string name, string typeName, bool hasDefault, string? defaultValue,
        bool isFileOrValue = false, string? fileSuffix = null,
        bool isFromString = false, string? exposedName = null,
        bool isRequired = false, bool isEnum = false,
        string? xmlDocDescription = null)
    {
        Name = name;
        TypeName = typeName;
        HasDefault = hasDefault;
        DefaultValue = defaultValue;
        IsFileOrValue = isFileOrValue;
        FileSuffix = fileSuffix;
        IsFromString = isFromString;
        ExposedName = exposedName;
        IsRequired = isRequired;
        IsEnum = isEnum;
        XmlDocDescription = xmlDocDescription;
    }
}

/// <summary>
/// Exposed parameter (aggregated across methods for CLI/MCP Settings).
/// </summary>
public sealed class ExposedParameter
{
    public string Name { get; }
    public string TypeName { get; }
    public string? Description { get; }
    public string? DefaultValue { get; }

    public ExposedParameter(string name, string typeName, string? description = null, string? defaultValue = null)
    {
        Name = name;
        TypeName = typeName;
        Description = description;
        DefaultValue = defaultValue;
    }
}
