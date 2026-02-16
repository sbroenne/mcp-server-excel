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
    public string? XmlDocSummary { get; }
    public List<MethodInfo> Methods { get; }

    /// <summary>Human-readable title for the MCP tool (e.g., "Excel Power Query Operations").</summary>
    public string? McpToolTitle { get; }

    /// <summary>Whether the tool is destructive (modifies data). Default: true.</summary>
    public bool McpToolDestructive { get; }

    /// <summary>MCP meta category (e.g., "data", "analysis", "query").</summary>
    public string? McpToolCategory { get; }

    /// <summary>Tool description for LLMs via [Description], since XML docs aren't available from metadata refs.</summary>
    public string? McpToolDescription { get; }

    /// <summary>Whether the interface has an explicit [McpTool] attribute. Used by MCP generator to skip hand-written tools.</summary>
    public bool HasMcpToolAttribute { get; }

    public ServiceInfo(string category, string categoryPascal, string mcpToolName, bool noSession, List<MethodInfo> methods,
        string? xmlDocSummary = null, string? mcpToolTitle = null, bool mcpToolDestructive = true, string? mcpToolCategory = null,
        string? mcpToolDescription = null, bool hasMcpToolAttribute = true)
    {
        Category = category;
        CategoryPascal = categoryPascal;
        McpToolName = mcpToolName;
        NoSession = noSession;
        XmlDocSummary = xmlDocSummary;
        Methods = methods;
        McpToolTitle = mcpToolTitle;
        McpToolDestructive = mcpToolDestructive;
        McpToolCategory = mcpToolCategory;
        McpToolDescription = mcpToolDescription;
        HasMcpToolAttribute = hasMcpToolAttribute;
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
    /// <summary>Whether the original interface method has an IExcelBatch parameter.</summary>
    public bool HasBatchParameter { get; }

    public MethodInfo(string methodName, string actionName, string returnType, string mcpTool,
        List<ParameterInfo> parameters, string? xmlDocSummary = null, bool hasBatchParameter = true)
    {
        MethodName = methodName;
        ActionName = actionName;
        ReturnType = returnType;
        McpTool = mcpTool;
        Parameters = parameters;
        XmlDocSummary = xmlDocSummary;
        HasBatchParameter = hasBatchParameter;
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

    /// <summary>
    /// The fully qualified enum type name when IsFromString and IsEnum are both true.
    /// Used by MCP generator to emit typed enum parameters instead of strings.
    /// </summary>
    public string? EnumTypeName { get; }

    public ParameterInfo(string name, string typeName, bool hasDefault, string? defaultValue,
        bool isFileOrValue = false, string? fileSuffix = null,
        bool isFromString = false, string? exposedName = null,
        bool isRequired = false, bool isEnum = false,
        string? xmlDocDescription = null, string? enumTypeName = null)
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
        EnumTypeName = enumTypeName;
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

    /// <summary>Action names where this parameter is required (non-nullable, no default, or [RequiredParameter]).</summary>
    public List<string> RequiredByActions { get; } = new();

    /// <summary>Total number of actions in the service (for computing "required for all" vs subset).</summary>
    public int TotalActionCount { get; set; }

    public ExposedParameter(string name, string typeName, string? description = null, string? defaultValue = null)
    {
        Name = name;
        TypeName = typeName;
        Description = description;
        DefaultValue = defaultValue;
    }

    /// <summary>
    /// Returns the description with required-by-actions suffix appended.
    /// E.g., "Name of the PivotTable (required for: read, delete, refresh)"
    /// or "Name of the PivotTable (required)" if required for all actions.
    /// </summary>
    public string? DescriptionWithRequired
    {
        get
        {
            if (RequiredByActions.Count == 0)
                return Description;

            var suffix = RequiredByActions.Count == TotalActionCount
                ? "(required)"
                : $"(required for: {string.Join(", ", RequiredByActions)})";

            return string.IsNullOrEmpty(Description)
                ? suffix
                : $"{Description} {suffix}";
        }
    }
}
