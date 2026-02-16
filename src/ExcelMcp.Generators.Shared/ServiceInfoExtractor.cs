using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Sbroenne.ExcelMcp.Generators.Common;

/// <summary>
/// Extracts ServiceInfo from interfaces marked with [ServiceCategory].
/// Shared between all generators.
/// </summary>
public static class ServiceInfoExtractor
{
    public static ServiceInfo? ExtractServiceInfo(INamedTypeSymbol interfaceSymbol)
    {
        string? category = null;
        string? pascalName = null;
        string? mcpTool = null;
        bool noSession = false;
        string? mcpToolTitle = null;
        bool mcpToolDestructive = true;
        string? mcpToolCategory = null;
        string? mcpToolDescription = null;

        foreach (var attr in interfaceSymbol.GetAttributes())
        {
            var attrName = attr.AttributeClass?.Name;

            if (attrName == "ServiceCategoryAttribute")
            {
                if (attr.ConstructorArguments.Length > 0)
                {
                    category = attr.ConstructorArguments[0].Value?.ToString();
                }
                if (attr.ConstructorArguments.Length > 1)
                {
                    pascalName = attr.ConstructorArguments[1].Value?.ToString();
                }
            }
            else if (attrName == "McpToolAttribute")
            {
                if (attr.ConstructorArguments.Length > 0)
                {
                    mcpTool = attr.ConstructorArguments[0].Value?.ToString();
                }

                // Read named properties: Title, Destructive, Category
                foreach (var namedArg in attr.NamedArguments)
                {
                    switch (namedArg.Key)
                    {
                        case "Title":
                            mcpToolTitle = namedArg.Value.Value?.ToString();
                            break;
                        case "Destructive":
                            if (namedArg.Value.Value is bool destructive)
                                mcpToolDestructive = destructive;
                            break;
                        case "Category":
                            mcpToolCategory = namedArg.Value.Value?.ToString();
                            break;
                        case "Description":
                            mcpToolDescription = namedArg.Value.Value?.ToString();
                            break;
                    }
                }
            }
            else if (attrName == "NoSessionAttribute")
            {
                noSession = true;
            }
        }

        if (category is null)
            return null;

        // Extract interface-level XML documentation
        var interfaceSummary = ExtractInterfaceSummary(interfaceSymbol);

        var methods = new List<MethodInfo>();

        foreach (var member in interfaceSymbol.GetMembers())
        {
            if (member is IMethodSymbol method && method.MethodKind == MethodKind.Ordinary)
            {
                var actionName = GetActionName(method);
                var methodMcpTool = GetMethodMcpTool(method) ?? mcpTool;
                var xmlDoc = ExtractXmlDocumentation(method);

                var hasBatchParameter = method.Parameters.Any(p => p.Type.Name == "IExcelBatch");
                var parameters = method.Parameters
                    .Where(p => p.Type.Name != "IExcelBatch") // Skip batch parameter
                    .Select(p => ExtractParameterInfo(p, xmlDoc))
                    .ToList();

                methods.Add(new MethodInfo(
                    method.Name,
                    actionName,
                    TypeNameHelper.GetTypeName(method.ReturnType),
                    methodMcpTool ?? category,
                    parameters,
                    xmlDoc?.Summary,
                    hasBatchParameter));
            }
        }

        // Use explicit pascalName if provided, otherwise derive from category
        var categoryPascal = pascalName ?? StringHelper.ToPascalCase(category);

        return new ServiceInfo(
            category,
            categoryPascal,
            mcpTool ?? category,
            noSession,
            methods,
            interfaceSummary,
            mcpToolTitle,
            mcpToolDestructive,
            mcpToolCategory,
            mcpToolDescription,
            hasMcpToolAttribute: mcpTool != null);
    }

    private static string? ExtractInterfaceSummary(INamedTypeSymbol interfaceSymbol)
    {
        var xmlComment = interfaceSymbol.GetDocumentationCommentXml();
        if (string.IsNullOrEmpty(xmlComment))
            return null;

        try
        {
            var doc = new XmlDocument();
            doc.LoadXml($"<root>{xmlComment}</root>");
            var text = doc.SelectSingleNode("//summary")?.InnerText?.Trim();
            if (string.IsNullOrEmpty(text))
                return null;

            // Normalize multi-line XML doc comments: collapse whitespace runs into single space
            return Regex.Replace(text, @"\s+", " ");
        }
        catch (Exception)
        {
            // XML comment parsing is best-effort — malformed XML should not break generation
            return null;
        }
    }

    private static string GetActionName(IMethodSymbol method)
    {
        // Check for [ServiceAction] override
        foreach (var attr in method.GetAttributes())
        {
            if (attr.AttributeClass?.Name == "ServiceActionAttribute" && attr.ConstructorArguments.Length > 0)
            {
                return attr.ConstructorArguments[0].Value?.ToString() ?? StringHelper.ToKebabCase(method.Name);
            }
        }

        // Default: derive from method name
        return StringHelper.ToKebabCase(method.Name);
    }

    private static string? GetMethodMcpTool(IMethodSymbol method)
    {
        foreach (var attr in method.GetAttributes())
        {
            if (attr.AttributeClass?.Name == "McpToolAttribute" && attr.ConstructorArguments.Length > 0)
            {
                return attr.ConstructorArguments[0].Value?.ToString();
            }
        }
        return null;
    }

    private static ParameterInfo ExtractParameterInfo(IParameterSymbol param, XmlDocumentation? methodDoc)
    {
        bool isFileOrValue = false;
        string? fileSuffix = null;
        bool isFromString = false;
        string? exposedName = null;
        bool isRequired = false;

        foreach (var attr in param.GetAttributes())
        {
            var attrName = attr.AttributeClass?.Name;

            if (attrName == "FileOrValueAttribute")
            {
                isFileOrValue = true;
                if (attr.ConstructorArguments.Length > 0)
                {
                    fileSuffix = attr.ConstructorArguments[0].Value?.ToString() ?? "File";
                }
                else
                {
                    fileSuffix = "File";
                }
            }
            else if (attrName == "FromStringAttribute")
            {
                isFromString = true;
                if (attr.ConstructorArguments.Length > 0)
                {
                    exposedName = attr.ConstructorArguments[0].Value?.ToString();
                }
            }
            else if (attrName == "RequiredParameterAttribute")
            {
                isRequired = true;
            }
        }

        // Detect if this is an enum type (including Nullable<Enum>)
        bool isEnum = param.Type.TypeKind == TypeKind.Enum;
        string? enumTypeName = null;
        if (isEnum)
        {
            enumTypeName = TypeNameHelper.GetTypeName(param.Type);
        }
        else if (param.Type is INamedTypeSymbol nullableType
            && nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T
            && nullableType.TypeArguments.Length == 1)
        {
            isEnum = nullableType.TypeArguments[0].TypeKind == TypeKind.Enum;
            if (isEnum)
            {
                enumTypeName = TypeNameHelper.GetTypeName(nullableType.TypeArguments[0]);
            }
        }

        // Get XML doc description for this parameter
        string? paramDescription = null;
        if (methodDoc?.Parameters != null && methodDoc.Parameters.TryGetValue(param.Name, out var desc))
        {
            paramDescription = desc;
        }

        return new ParameterInfo(
            param.Name,
            TypeNameHelper.GetTypeName(param.Type, param.NullableAnnotation),
            param.HasExplicitDefaultValue,
            param.HasExplicitDefaultValue ? TypeNameHelper.GetDefaultValueString(param) : null,
            isFileOrValue,
            fileSuffix,
            isFromString,
            exposedName,
            isRequired,
            isEnum,
            paramDescription,
            enumTypeName);
    }

    private static XmlDocumentation? ExtractXmlDocumentation(IMethodSymbol method)
    {
        var xmlComment = method.GetDocumentationCommentXml();
        if (string.IsNullOrEmpty(xmlComment))
            return null;

        try
        {
            var doc = new XmlDocument();
            doc.LoadXml($"<root>{xmlComment}</root>");

            var summary = doc.SelectSingleNode("//summary")?.InnerText?.Trim();
            var parameters = new Dictionary<string, string>();

            var paramNodes = doc.SelectNodes("//param");
            if (paramNodes != null)
            {
                foreach (XmlNode paramNode in paramNodes)
                {
                    var name = paramNode.Attributes?["name"]?.Value;
                    var description = paramNode.InnerText?.Trim();
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(description))
                    {
                        parameters[name] = description;
                    }
                }
            }

            return new XmlDocumentation(summary, parameters);
        }
        catch (Exception)
        {
            // XML doc parsing is best-effort — malformed XML should not break generation
            return null;
        }
    }

    /// <summary>
    /// Gets all unique exposed parameters across all methods in a service.
    /// Tracks which actions require each parameter for description enrichment.
    /// </summary>
    public static List<ExposedParameter> GetAllExposedParameters(ServiceInfo info)
    {
        var paramMap = new Dictionary<string, ExposedParameter>(StringComparer.OrdinalIgnoreCase);

        foreach (var method in info.Methods)
        {
            foreach (var p in method.Parameters)
            {
                // Get the exposed name (from attribute or original name)
                var exposedName = p.ExposedName ?? p.Name;

                if (!paramMap.TryGetValue(exposedName, out var existing))
                {
                    existing = new ExposedParameter(exposedName, p.TypeName, p.XmlDocDescription);
                    paramMap[exposedName] = existing;
                }

                // Track if this param is required for this action
                var isRequired = p.IsRequired || (!p.HasDefault && !p.TypeName.EndsWith("?"));
                if (isRequired)
                {
                    existing.RequiredByActions.Add(method.ActionName);
                }

                // If FileOrValue, also add the file variant
                if (p.IsFileOrValue && p.FileSuffix != null)
                {
                    var fileParamName = exposedName + p.FileSuffix;
                    if (!paramMap.ContainsKey(fileParamName))
                    {
                        paramMap[fileParamName] = new ExposedParameter(fileParamName, "string?", $"Path to file containing {exposedName}");
                    }
                }
            }
        }

        // Set total action count on all params
        var totalActions = info.Methods.Count;
        foreach (var ep in paramMap.Values)
        {
            ep.TotalActionCount = totalActions;
        }

        return paramMap.Values.ToList();
    }
}

/// <summary>
/// Extracted XML documentation from a method.
/// </summary>
public sealed class XmlDocumentation
{
    public string? Summary { get; }
    public Dictionary<string, string> Parameters { get; }

    public XmlDocumentation(string? summary, Dictionary<string, string> parameters)
    {
        Summary = summary;
        Parameters = parameters;
    }
}
