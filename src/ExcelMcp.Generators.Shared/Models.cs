using Microsoft.CodeAnalysis;

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
    public MacCapabilityInfo MacCapability { get; }

    public ServiceInfo(string category, string categoryPascal, string mcpToolName, bool noSession, List<MethodInfo> methods,
        string? xmlDocSummary = null, string? mcpToolTitle = null, bool mcpToolDestructive = true, string? mcpToolCategory = null,
        string? mcpToolDescription = null, bool hasMcpToolAttribute = true, MacCapabilityInfo? macCapability = null)
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
        MacCapability = macCapability ?? MacCapabilityInfo.Unclassified;
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
    public ITypeSymbol ReturnTypeSymbol { get; }
    public string McpTool { get; }
    public List<ParameterInfo> Parameters { get; }
    public string? XmlDocSummary { get; }
    /// <summary>Whether the original interface method has an IExcelBatch parameter.</summary>
    public bool HasBatchParameter { get; }

    /// <summary>Whether the original interface method has an IProgress&lt;T&gt; parameter.</summary>
    public bool HasProgressParameter { get; }
    public MacCapabilityInfo MacCapability { get; }

    public MethodInfo(string methodName, string actionName, string returnType, ITypeSymbol returnTypeSymbol, string mcpTool,
        List<ParameterInfo> parameters, string? xmlDocSummary = null, bool hasBatchParameter = true,
        bool hasProgressParameter = false, MacCapabilityInfo? macCapability = null)
    {
        MethodName = methodName;
        ActionName = actionName;
        ReturnType = returnType;
        ReturnTypeSymbol = returnTypeSymbol;
        McpTool = mcpTool;
        Parameters = parameters;
        XmlDocSummary = xmlDocSummary;
        HasBatchParameter = hasBatchParameter;
        HasProgressParameter = hasProgressParameter;
        MacCapability = macCapability ?? MacCapabilityInfo.Unclassified;
    }
}

public sealed class MacCapabilityInfo
{
    public static MacCapabilityInfo Unclassified { get; } = new(
        "Unsupported",
        "Blocked",
        false,
        "The Apple Events and Office.js catalogs expose no action-specific route proven to preserve this generated Windows contract.",
        "Excel for Mac 16.113.1 Apple Events dictionary; Office.js ExcelApi requirement-set review.",
        "current supported macOS APIs cannot preserve the exact public contract; use the Windows COM backend");

    public MacCapabilityInfo(
        string tier,
        string status,
        bool isAvailable,
        string evidence,
        string excelApiVersion,
        string blocker)
    {
        Tier = tier;
        Status = status;
        IsAvailable = isAvailable;
        Evidence = evidence;
        ExcelApiVersion = excelApiVersion;
        Blocker = blocker;
    }

    public string Tier { get; }
    public string Status { get; }
    public bool IsAvailable { get; }
    public string Evidence { get; }
    public string ExcelApiVersion { get; }
    public string Blocker { get; }
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
    public bool AllowsEmptyString { get; }
    public bool IsParams { get; }
    public bool IsEnum { get; }
    public bool IsJsonObject { get; }
    public string? XmlDocDescription { get; }

    /// <summary>
    /// The fully qualified enum type name when IsFromString and IsEnum are both true.
    /// Used by MCP generator to emit typed enum parameters instead of strings.
    /// </summary>
    public string? EnumTypeName { get; }
    public IReadOnlyList<EnumAliasInfo> EnumAliases { get; }

    public ParameterInfo(string name, string typeName, bool hasDefault, string? defaultValue,
        bool isFileOrValue = false, string? fileSuffix = null,
        bool isFromString = false, string? exposedName = null,
        bool isRequired = false, bool isEnum = false,
        string? xmlDocDescription = null, string? enumTypeName = null,
        IReadOnlyList<EnumAliasInfo>? enumAliases = null, bool isParams = false,
        bool allowsEmptyString = false, bool isJsonObject = false)
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
        AllowsEmptyString = allowsEmptyString;
        IsParams = isParams;
        IsEnum = isEnum;
        IsJsonObject = isJsonObject;
        XmlDocDescription = xmlDocDescription;
        EnumTypeName = enumTypeName;
        EnumAliases = enumAliases ?? [];
    }
}

public sealed class EnumAliasInfo
{
    public EnumAliasInfo(string alias, string memberName)
    {
        Alias = alias;
        MemberName = memberName;
    }

    public string Alias { get; }
    public string MemberName { get; }
}

/// <summary>
/// Exposed parameter (aggregated across methods for CLI/MCP Settings).
/// </summary>
public sealed class ExposedParameter
{
    public string Name { get; }
    public string TypeName { get; set; }
    public string? Description { get; }
    public string? DefaultValue { get; }

    /// <summary>Action names where this parameter is required (non-nullable, no default, or [RequiredParameter]).</summary>
    public List<string> RequiredByActions { get; } = new();

    /// <summary>Total number of actions in the service (for computing "required for all" vs subset).</summary>
    public int TotalActionCount { get; set; }
    public List<string> ApplicableByActions { get; } = new();

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
            var suffixes = new List<string>();
            if (RequiredByActions.Count > 0)
            {
                suffixes.Add(RequiredByActions.Count == TotalActionCount
                    ? "(required)"
                    : $"(required for: {string.Join(", ", RequiredByActions)})");
            }
            if (ApplicableByActions.Count > 0 && ApplicableByActions.Count < TotalActionCount)
            {
                suffixes.Add($"(valid for: {string.Join(", ", ApplicableByActions)})");
            }

            if (suffixes.Count == 0)
                return Description;

            var suffix = string.Join(" ", suffixes);
            return string.IsNullOrEmpty(Description) ? suffix : $"{Description} {suffix}";
        }
    }
}
