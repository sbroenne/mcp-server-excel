using System.Text;
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
            else if (attrName == "McpToolAttribute" && attr.ConstructorArguments.Length > 0)
            {
                mcpTool = attr.ConstructorArguments[0].Value?.ToString();
            }
            else if (attrName == "NoSessionAttribute")
            {
                noSession = true;
            }
        }

        if (category is null)
            return null;

        var methods = new List<MethodInfo>();

        foreach (var member in interfaceSymbol.GetMembers())
        {
            if (member is IMethodSymbol method && method.MethodKind == MethodKind.Ordinary)
            {
                var actionName = GetActionName(method);
                var methodMcpTool = GetMethodMcpTool(method) ?? mcpTool;
                var xmlDoc = ExtractXmlDocumentation(method);

                var parameters = method.Parameters
                    .Where(p => p.Type.Name != "IExcelBatch") // Skip batch parameter
                    .Select(p => ExtractParameterInfo(p, xmlDoc))
                    .ToList();

                methods.Add(new MethodInfo(
                    method.Name,
                    actionName,
                    TypeNameHelper.GetTypeName(method.ReturnType),
                    methodMcpTool ?? "unknown",
                    parameters,
                    xmlDoc?.Summary));
            }
        }

        // Use explicit pascalName if provided, otherwise derive from category
        var categoryPascal = pascalName ?? StringHelper.ToPascalCase(category);

        return new ServiceInfo(
            category,
            categoryPascal,
            mcpTool ?? "unknown",
            noSession,
            methods);
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

        // Detect if this is an enum type
        bool isEnum = param.Type.TypeKind == TypeKind.Enum;

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
            paramDescription);
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
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Gets all unique exposed parameters across all methods in a service.
    /// </summary>
    public static List<ExposedParameter> GetAllExposedParameters(ServiceInfo info)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<ExposedParameter>();

        foreach (var method in info.Methods)
        {
            foreach (var p in method.Parameters)
            {
                // Get the exposed name (from attribute or original name)
                var exposedName = p.ExposedName ?? p.Name;
                if (!seen.Add(exposedName))
                    continue;

                result.Add(new ExposedParameter(exposedName, p.TypeName, p.XmlDocDescription));

                // If FileOrValue, also add the file variant
                if (p.IsFileOrValue && p.FileSuffix != null)
                {
                    var fileParamName = exposedName + p.FileSuffix;
                    if (seen.Add(fileParamName))
                    {
                        result.Add(new ExposedParameter(fileParamName, "string?", $"Path to file containing {exposedName}"));
                    }
                }
            }
        }

        return result;
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
