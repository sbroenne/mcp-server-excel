using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Sbroenne.ExcelMcp.Generators.Common;

namespace Sbroenne.ExcelMcp.Generators.Mcp;

/// <summary>
/// Generates MCP Server tool classes from Core [ServiceCategory] interfaces.
/// Discovers interfaces from referenced assemblies (Core) and generates
/// [McpServerToolType] classes with properly typed parameters.
///
/// Key features:
/// - Tool action stays a typed enum for discoverable required actions
/// - Optional [FromString] enum parameters stay strings to avoid nullable-enum schema sentinels
/// - TimeSpan parameters become int (seconds) for JSON compatibility
/// - FileOrValue parameters generate dual parameters (value + file path)
/// - XML docs from Core interfaces become MCP tool descriptions
/// </summary>
[Generator]
public class McpToolGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var documentation = context.AdditionalTextsProvider
            .Where(file => file.Path.EndsWith("Sbroenne.ExcelMcp.Core.xml", StringComparison.OrdinalIgnoreCase))
            .Select((file, cancellationToken) => file.GetText(cancellationToken)?.ToString())
            .Collect();

        context.RegisterSourceOutput(context.CompilationProvider.Combine(documentation),
            static (spc, input) =>
            {
                var compilation = input.Left;
                var coreReference = compilation.References.OfType<PortableExecutableReference>()
                    .SingleOrDefault(reference =>
                        compilation.GetAssemblyOrModuleSymbol(reference)?.Name == "Sbroenne.ExcelMcp.Core");
                if (coreReference != null)
                {
                    if (input.Right.Length != 1 || string.IsNullOrWhiteSpace(input.Right[0]))
                    {
                        spc.ReportDiagnostic(Diagnostic.Create(
                            new DiagnosticDescriptor("EXCELMCP001", "Missing Core documentation",
                                "MCP generation requires the Core XML documentation as an AdditionalFile",
                                "ExcelMcp.Generation", DiagnosticSeverity.Error, isEnabledByDefault: true),
                            Location.None));
                        return;
                    }

                    compilation = compilation.ReplaceReference(coreReference,
                        ((AssemblyMetadata)coreReference.GetMetadata()).GetReference(
                            documentation: new CoreDocumentationProvider(input.Right[0]!),
                            aliases: coreReference.Properties.Aliases,
                            embedInteropTypes: coreReference.Properties.EmbedInteropTypes,
                            filePath: coreReference.FilePath));
                }

                var services = DiscoverServices(compilation);
                if (services.Count == 0)
                    return;

                foreach (var info in services.SelectMany(SplitReadOnlyTools))
                {
                    var code = GenerateToolClass(info);
                    var suffix = info.McpToolReadOnly ? ".ReadOnly" : string.Empty;
                    spc.AddSource($"McpTool.{info.CategoryPascal}{suffix}.g.cs", SourceText.From(code, Encoding.UTF8));
                }
            });
    }

    /// <summary>
    /// Discovers [ServiceCategory] interfaces from referenced assemblies and extracts ServiceInfo.
    /// Skips generation for any tool name that already has a manual [McpServerTool] implementation
    /// in the current compilation, preventing duplicate tool registration.
    /// </summary>
    private static List<ServiceInfo> DiscoverServices(Compilation compilation)
    {
        var manualToolNames = DiscoverManualToolNames(compilation);
        var result = new List<ServiceInfo>();

        foreach (var reference in compilation.References)
        {
            if (compilation.GetAssemblyOrModuleSymbol(reference) is not IAssemblySymbol assembly)
                continue;

            foreach (var type in GetAllTypes(assembly.GlobalNamespace))
            {
                if (type.TypeKind != TypeKind.Interface)
                    continue;

                var info = ServiceInfoExtractor.ExtractServiceInfo(type);
                if (info != null && info.HasMcpToolAttribute && !manualToolNames.Contains(info.McpToolName))
                    result.Add(info);
            }
        }

        return result;
    }

    /// <summary>
    /// Discovers tool names that are manually defined in the current compilation via
    /// [McpServerTool(Name = "...")] attributes on methods. Used to avoid generating
    /// tools that would conflict with manual implementations.
    /// </summary>
    private static HashSet<string> DiscoverManualToolNames(Compilation compilation)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var type in GetAllTypes(compilation.Assembly.GlobalNamespace))
        {
            foreach (var member in type.GetMembers())
            {
                if (member is not IMethodSymbol method)
                    continue;

                foreach (var attr in method.GetAttributes())
                {
                    if (attr.AttributeClass?.Name != "McpServerToolAttribute")
                        continue;

                    foreach (var namedArg in attr.NamedArguments)
                    {
                        if (namedArg.Key == "Name" && namedArg.Value.Value is string name)
                        {
                            result.Add(name);
                        }
                    }
                }
            }
        }

        return result;
    }

    private static IEnumerable<INamedTypeSymbol> GetAllTypes(INamespaceSymbol ns)
    {
        foreach (var type in ns.GetTypeMembers())
            yield return type;
        foreach (var child in ns.GetNamespaceMembers())
            foreach (var type in GetAllTypes(child))
                yield return type;
    }

    private static IEnumerable<ServiceInfo> SplitReadOnlyTools(ServiceInfo info)
    {
        var readOnlyMethods = info.Methods.Where(method => method.McpToolReadOnly).ToArray();
        var readOnlyActions = readOnlyMethods.Select(method => method.ActionName).ToArray();
        var readOnlyToolName = readOnlyMethods.FirstOrDefault()?.McpTool;

        foreach (var group in info.Methods.GroupBy(method => method.McpTool, StringComparer.Ordinal))
        {
            var methods = group.ToList();
            var readOnly = methods.All(method => method.McpToolReadOnly);
            var description = readOnly
                ? "Read-only actions: " + string.Join(" ", methods.Select(method =>
                    string.IsNullOrWhiteSpace(method.XmlDocSummary)
                        ? $"{method.ActionName}."
                        : $"{method.ActionName}: {method.XmlDocSummary}"))
                : readOnlyActions.Length > 0 && readOnlyToolName is not null
                    ? $"{info.McpToolDescription} Read-only actions {string.Join(", ", readOnlyActions)} are available through {readOnlyToolName}."
                    : info.McpToolDescription;
            yield return new ServiceInfo(
                info.Category,
                info.CategoryPascal,
                group.Key,
                info.NoSession,
                methods,
                readOnly ? $"Read-only {info.CategoryPascal} actions." : info.XmlDocSummary,
                readOnly ? $"Read-only {info.CategoryPascal} Operations" : info.McpToolTitle,
                readOnly ? false : info.McpToolDestructive,
                readOnly,
                info.McpToolCategory,
                description,
                info.HasMcpToolAttribute);
        }
    }

    /// <summary>
    /// Generates a complete MCP tool class for a service category.
    /// </summary>
    private static string GenerateToolClass(ServiceInfo info)
    {
        var sb = new StringBuilder();
        var hasProgress = info.Methods.Any(m => m.HasProgressParameter);

        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("// Generator version: nullable-fix-v2");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member");
        sb.AppendLine();
        sb.AppendLine("using System.ComponentModel;");
        sb.AppendLine("using System.Text.Json.Serialization;");
        sb.AppendLine("using System.Threading;");
        sb.AppendLine("using ModelContextProtocol.Protocol;");
        sb.AppendLine("using ModelContextProtocol.Server;");
        sb.AppendLine("using Sbroenne.ExcelMcp.Generated;");
        if (hasProgress)
        {
            sb.AppendLine("using ModelContextProtocol;");
            sb.AppendLine("using Sbroenne.ExcelMcp.ComInterop;");
            sb.AppendLine("using Sbroenne.ExcelMcp.McpServer.Progress;");
        }
        sb.AppendLine();
        sb.AppendLine("namespace Sbroenne.ExcelMcp.McpServer.Tools;");
        sb.AppendLine();

        // Class XML doc
        sb.AppendLine("/// <summary>");
        sb.AppendLine($"/// Generated MCP tool for {info.CategoryPascal} operations.");
        sb.AppendLine("/// </summary>");
        sb.AppendLine("[McpServerToolType]");

        var className = GetClassName(info);
        sb.AppendLine($"public static partial class {className}");
        sb.AppendLine("{");

        // Generate the tool method
        GenerateActionEnum(sb, info);
        GenerateToolMethod(sb, info, hasProgress);

        sb.AppendLine("}");
        sb.AppendLine();
        GenerateOutputSchemaClass(sb, info);
        return sb.ToString();
    }

    private static void GenerateActionEnum(StringBuilder sb, ServiceInfo info)
    {
        var enumTypeName = GetActionTypeName(info);
        sb.AppendLine("    [JsonConverter(typeof(JsonStringEnumConverter<" + enumTypeName + ">))]");
        sb.AppendLine($"    public enum {enumTypeName}");
        sb.AppendLine("    {");
        for (var i = 0; i < info.Methods.Count; i++)
        {
            var method = info.Methods[i];
            var comma = i < info.Methods.Count - 1 ? "," : "";
            sb.AppendLine($"        [JsonStringEnumMemberName(\"{method.ActionName}\")]");
            sb.AppendLine($"        {method.MethodName}{comma}");
        }
        sb.AppendLine("    }");
        sb.AppendLine();
    }

    /// <summary>
    /// Generates the MCP tool method with XML docs, attributes, parameters, and body.
    /// </summary>
    private static void GenerateToolMethod(StringBuilder sb, ServiceInfo info, bool hasProgress)
    {
        var enumTypeName = GetActionTypeName(info);

        // Get all exposed parameters (aggregated across methods)
        var exposedParams = ServiceInfoExtractor.GetAllExposedParameters(info);

        // Build the enhanced parameter list for MCP
        var mcpParams = BuildMcpParameters(info, exposedParams);

        // XML doc: interface-level summary
        var summary = info.McpToolReadOnly ? $"Read-only {info.CategoryPascal} actions." : info.XmlDocSummary;
        if (!string.IsNullOrEmpty(summary))
        {
            sb.AppendLine("    /// <summary>");
            // Wrap the summary text, respecting line breaks
            foreach (var line in WrapXmlDocLines(summary))
            {
                sb.AppendLine($"    /// {EscapeXml(line)}");
            }
            sb.AppendLine("    /// </summary>");
        }

        // XML doc: parameter descriptions
        sb.AppendLine($"    /// <param name=\"action\">The action to perform</param>");
        if (!info.NoSession)
        {
            sb.AppendLine($"    /// <param name=\"session_id\">Session ID from file 'open' action</param>");
        }
        foreach (var p in mcpParams)
        {
            if (!string.IsNullOrEmpty(p.Description))
            {
                var escapedDesc = EscapeXml(p.Description);
                sb.AppendLine($"    /// <param name=\"{p.Name}\">{escapedDesc}</param>");
            }
        }

        // Attributes
        var title = info.McpToolTitle ?? $"Excel {info.CategoryPascal} Operations";
        var destructive = info.McpToolDestructive ? "true" : "false";
        var readOnly = info.McpToolReadOnly ? "true" : "false";
        sb.AppendLine($"    [McpServerTool(Name = \"{info.McpToolName}\", Title = \"{title}\", Destructive = {destructive}, ReadOnly = {readOnly}, UseStructuredContent = true, OutputSchemaType = typeof({GetOutputSchemaClassName(info)}))]");

        var category = info.McpToolCategory ?? "data";
        sb.AppendLine($"    [McpMeta(\"category\", \"{category}\")]");
        sb.AppendLine($"    [McpMeta(\"requiresSession\", {(!info.NoSession).ToString().ToLower()})]");

        // Tool prose stays concise; parameter details come from Core XML documentation.
        if (!string.IsNullOrEmpty(info.McpToolDescription))
        {
            var methodDesc = EscapeStringLiteral(RenderMcpParameterNames(info.McpToolDescription, exposedParams));
            sb.AppendLine($"    [Description(\"{methodDesc}\")]");
        }

        // Method signature — non-partial because MCP SDK's XmlToDescriptionGenerator
        // cannot see our generator output to create a matching defining declaration.
        var methodName = GetMethodName(info);
        sb.Append($"    public static async Task<CallToolResult> {methodName}(");
        sb.AppendLine();

        // Keep action as a required enum so schema consumers get a strict enum without nullable sentinels.
        sb.AppendLine($"        [Description(\"The action to perform\")] {enumTypeName} action,");
        sb.AppendLine("        Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge bridge,");

        // Session parameter (if required)
        if (!info.NoSession)
        {
            sb.Append("        [Description(\"Session ID from file 'open' action\")] string session_id");
            sb.Append(",");
            sb.AppendLine();
        }

        // Exposed parameters are always optional at the MCP tool surface because each action
        // uses a different subset. Emit C# optional defaults instead of [DefaultValue] attributes
        // so nullable schema parameters stay optional without SDK-inserted enum sentinels.
        for (int i = 0; i < mcpParams.Count; i++)
        {
            var p = mcpParams[i];
            var defaultExpr = p.DefaultExpression ?? "null";
            if (!string.IsNullOrEmpty(p.Description))
            {
                var desc = EscapeStringLiteral(p.Description);
                sb.Append($"        [Description(\"{desc}\")] {p.McpTypeName} {p.Name} = {defaultExpr}");
            }
            else
            {
                sb.Append($"        {p.McpTypeName} {p.Name} = {defaultExpr}");
            }
            sb.Append(",");
            sb.AppendLine();
        }

        // DI-injected progress parameter (no [Description] — resolved from RequestServiceProvider)
        if (hasProgress)
        {
            sb.AppendLine("        IProgress<ProgressNotificationValue> progress = default!,");
        }
        sb.AppendLine("        RequestContext<CallToolRequestParams> requestContext = default!,");
        sb.AppendLine("        CancellationToken cancellationToken = default");
        sb.AppendLine("    )");
        sb.AppendLine("    {");

        // Method body: pre-processing and RouteAction call
        GenerateMethodBody(sb, info, mcpParams, enumTypeName, hasProgress);

        sb.AppendLine("    }");
    }

    private static void GenerateOutputSchemaClass(StringBuilder sb, ServiceInfo info)
    {
        sb.AppendLine($"internal sealed class {GetOutputSchemaClassName(info)}");
        sb.AppendLine("{");

        foreach (var property in GetOutputSchemaProperties(info))
        {
            if (property.Name == "SessionId")
                sb.AppendLine("    [JsonPropertyName(\"session_id\")]");
            sb.AppendLine("    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]");
            sb.AppendLine($"    public {property.TypeName} {property.Name} {{ get; set; }}");
        }

        sb.AppendLine("}");
    }

    private static string GetOutputSchemaClassName(ServiceInfo info) =>
        $"{info.CategoryPascal}{(info.McpToolReadOnly ? "ReadOnly" : "")}ToolOutputSchema";

    private static OutputSchemaProperty[] GetOutputSchemaProperties(ServiceInfo info)
    {
        var properties = new Dictionary<string, OutputSchemaProperty>(StringComparer.Ordinal);
        properties.Add("SessionId", new("SessionId", "string?"));

        foreach (var method in info.Methods)
        {
            if (method.ReturnTypeSymbol is not INamedTypeSymbol returnType || !DerivesFromResultBase(returnType))
            {
                if (method.ReturnTypeSymbol.SpecialType != SpecialType.System_Void)
                {
                    var typeName = GetOptionalSchemaTypeName(method.ReturnTypeSymbol);
                    if (properties.TryGetValue("Result", out var existing) && existing.TypeName != typeName)
                    {
                        properties["Result"] = new("Result", "System.Text.Json.JsonElement?");
                    }
                    else
                    {
                        properties["Result"] = new("Result", typeName);
                    }
                }

                continue;
            }

            for (var type = returnType; type is not null; type = type.BaseType)
            {
                foreach (var property in type.GetMembers().OfType<IPropertySymbol>())
                {
                    if (property.IsStatic || property.IsIndexer ||
                        property.DeclaredAccessibility != Accessibility.Public || property.GetMethod is null)
                        continue;

                    var typeName = GetOptionalSchemaTypeName(property.Type);
                    if (properties.TryGetValue(property.Name, out var existing) && existing.TypeName != typeName)
                    {
                        properties[property.Name] = new(property.Name, "System.Text.Json.JsonElement?");
                    }
                    else
                    {
                        properties[property.Name] = new(property.Name, typeName);
                    }
                }
            }
        }

        return properties.Values
            .OrderBy(p => p.Name == "Success" ? 0 : 1)
            .ThenBy(p => p.Name, StringComparer.Ordinal)
            .ToArray();
    }

    private static string GetOptionalSchemaTypeName(ITypeSymbol type)
    {
        var typeName = TypeNameHelper.GetTypeName(type, type.NullableAnnotation);
        if (!typeName.EndsWith("?", StringComparison.Ordinal))
            typeName += "?";

        return typeName;
    }

    private static bool DerivesFromResultBase(INamedTypeSymbol type)
    {
        for (var current = type; current is not null; current = current.BaseType)
        {
            if (current.Name == "ResultBase" &&
                current.ContainingNamespace.ToDisplayString() == "Sbroenne.ExcelMcp.Core.Models")
            {
                return true;
            }
        }

        return false;
    }

    private sealed class OutputSchemaProperty(string name, string typeName)
    {
        public string Name { get; } = name;
        public string TypeName { get; } = typeName;
    }

    /// <summary>
    /// Generates the method body that calls ServiceRegistry.RouteAction.
    /// </summary>
    private static void GenerateMethodBody(StringBuilder sb, ServiceInfo info, List<McpParameter> mcpParams, string enumTypeName, bool hasProgress)
    {
        var registryName = info.CategoryPascal;
        var toolName = info.McpToolName;

        // Set ambient progress context so DispatchToCore can inject it into Core methods
        if (hasProgress)
        {
            sb.AppendLine("        var previousProgress = ProgressContext.Current;");
            sb.AppendLine("        ProgressContext.Current = new McpProgressAdapter(progress);");
            sb.AppendLine("        try");
            sb.AppendLine("        {");
        }

        var indent = hasProgress ? "            " : "        ";

        sb.AppendLine($"{indent}var serviceAction = action switch");
        sb.AppendLine($"{indent}{{");
        foreach (var method in info.Methods)
            sb.AppendLine($"{indent}    {enumTypeName}.{method.MethodName} => {registryName}Action.{method.MethodName},");
        sb.AppendLine($"{indent}    _ => throw new ArgumentException($\"Unknown {enumTypeName}: {{action}}\")");
        sb.AppendLine($"{indent}}};");
        sb.AppendLine();
        const string serviceAction = "serviceAction";

        sb.AppendLine($"{indent}return await ExcelToolsBase.ExecuteToolActionAsync(");
        sb.AppendLine($"{indent}    \"{toolName}\",");
        sb.AppendLine($"{indent}    ServiceRegistry.{registryName}.ToActionString({serviceAction}),");
        sb.AppendLine($"{indent}    () =>");
        sb.AppendLine($"{indent}    {{");
        if (mcpParams.Count > 0)
        {
            sb.AppendLine($"{indent}        ServiceRegistry.{registryName}.ValidateActionParameters(");
            sb.AppendLine($"{indent}            ServiceRegistry.{registryName}.ToActionString({serviceAction}),");
            sb.AppendLine($"{indent}            ServiceRegistry.GetSuppliedParameterNames(");
            for (int i = 0; i < mcpParams.Count; i++)
            {
                var parameter = mcpParams[i];
                var comma = i < mcpParams.Count - 1 ? "," : "),";
                sb.AppendLine(
                    $"{indent}                (\"{parameter.RouteActionParamName}\", requestContext.Params.Arguments?.ContainsKey(\"{parameter.Name}\") == true ? true : (bool?)null){comma}");
            }
            sb.AppendLine($"{indent}            allowFileParameters: true);");
            sb.AppendLine();
        }
        foreach (var parameter in mcpParams.Where(parameter => parameter.PreProcessingCode != null))
        {
            sb.AppendLine($"{indent}        {parameter.PreProcessingCode}");
        }
        if (mcpParams.Any(parameter => parameter.PreProcessingCode != null))
        {
            sb.AppendLine();
        }
        sb.AppendLine($"{indent}        return ServiceRegistry.{registryName}.RouteAction(");
        sb.AppendLine($"{indent}            {serviceAction},");

        if (!info.NoSession)
        {
            sb.AppendLine($"{indent}            session_id,");
        }
        else
        {
            sb.AppendLine($"{indent}            \"\",");
        }

        var routeCallbackComma = mcpParams.Count > 0 ? "," : string.Empty;
        sb.AppendLine($"{indent}            (command, id, args) => ExcelToolsBase.ForwardToServiceAsync(bridge, command, id, args, cancellationToken){routeCallbackComma}");

        // Named arguments to RouteAction
        for (int i = 0; i < mcpParams.Count; i++)
        {
            var p = mcpParams[i];
            var routeArgName = p.RouteActionParamName;
            var routeArgValue = p.RouteActionValue;
            var comma = i < mcpParams.Count - 1 ? "," : "";
            sb.AppendLine($"{indent}            {routeArgName}: {routeArgValue}{comma}");
        }

        sb.AppendLine($"{indent}        );");
        sb.AppendLine($"{indent}    }}, cancellationToken);");

        // Close progress try/finally block
        if (hasProgress)
        {
            sb.AppendLine("        }");
            sb.AppendLine("        finally");
            sb.AppendLine("        {");
            sb.AppendLine("            ProgressContext.Current = previousProgress;");
            sb.AppendLine("        }");
        }
    }

    /// <summary>
    /// Builds MCP parameter descriptors from the exposed parameters.
    /// Handles type conversions: optional FromString enums stay strings, TimeSpan → int seconds, etc.
    /// </summary>
    private static List<McpParameter> BuildMcpParameters(ServiceInfo info, List<ExposedParameter> exposedParams)
    {
        var result = new List<McpParameter>();

        // Build a lookup of param info by exposed name for type resolution
        var paramInfoByName = new Dictionary<string, ParameterInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var method in info.Methods)
        {
            foreach (var p in method.Parameters)
            {
                var exposedName = p.ExposedName ?? p.Name;
                if (!paramInfoByName.ContainsKey(exposedName))
                    paramInfoByName[exposedName] = p;
            }
        }

        foreach (var ep in exposedParams)
        {
            paramInfoByName.TryGetValue(ep.Name, out var pInfo);
            var snakeName = StringHelper.ToSnakeCase(ep.Name);
            var description = ep.DescriptionWithRequired;
            if (string.IsNullOrWhiteSpace(description) ||
                description.StartsWith("(required", StringComparison.Ordinal) ||
                description.StartsWith("(valid", StringComparison.Ordinal))
            {
                var actionContext = BuildParameterActionContext(info, ep);
                if (!string.IsNullOrWhiteSpace(actionContext))
                {
                    description = string.Join(" ", new[] { actionContext, description }
                        .Where(value => !string.IsNullOrWhiteSpace(value)));
                }
            }
            if (description is not null)
                description = RenderMcpParameterNames(description, exposedParams);

            // Determine MCP type and conversion
            if (pInfo != null && pInfo.IsEnum && pInfo.EnumTypeName != null)
            {
                if (pInfo.IsFromString)
                {
                    // ServiceRegistry owns the shared enum aliases and case-insensitive parsing.
                    result.Add(new McpParameter(
                        name: snakeName,
                        mcpTypeName: "string?",
                        routeActionParamName: ep.Name,
                        routeActionValue: snakeName,
                        description: description,
                        defaultExpression: "null",
                        preProcessingCode: null));
                }
                else
                {
                    var localVarName = $"_{ep.Name}Parsed";
                    var aliasArguments = BuildEnumAliasArguments(pInfo);
                    var preProcessingCode =
                        $"var {localVarName} = !string.IsNullOrEmpty({snakeName}) ? ({pInfo.EnumTypeName}?)ServiceRegistry.ParseEnumValue<{pInfo.EnumTypeName}>({snakeName}, default, \"{ep.Name}\"{aliasArguments}) : null;";

                    result.Add(new McpParameter(
                        name: snakeName,
                        mcpTypeName: "string?",
                        routeActionParamName: ep.Name,
                        routeActionValue: localVarName,
                        description: description,
                        defaultExpression: "null",
                        preProcessingCode: preProcessingCode));
                }
            }
            else if (ep.TypeName.Contains("TimeSpan"))
            {
                // Public timeout inputs are whole seconds. Conversion happens once in service dispatch.
                var secondsName = ep.Name + "Seconds";
                var snakeSecondsName = StringHelper.ToSnakeCase(secondsName);
                var minimumSeconds = info.Category == "powerquery" ? 0 : 1;
                result.Add(new McpParameter(
                    name: snakeSecondsName,
                    mcpTypeName: "int?",
                    routeActionParamName: ep.Name,
                    routeActionValue: snakeSecondsName,
                    description: description != null
                        ? description + $" Accepted range: {minimumSeconds}-2147483 seconds."
                        : $"Timeout in whole seconds; range {minimumSeconds}-2147483",
                    defaultExpression: "null",
                    preProcessingCode: null));
            }
            else if (ep.TypeName.StartsWith("System.Collections.Generic.List<string>") ||
                     ep.TypeName.StartsWith("List<string>"))
            {
                // List<string> → string (JSON array) in MCP, parse via ParseJsonList
                var localVarName = $"_{ep.Name}Parsed";
                result.Add(new McpParameter(
                    name: snakeName,
                    mcpTypeName: "string?",
                    routeActionParamName: ep.Name,
                    routeActionValue: localVarName,
                    description: description != null
                        ? description + " (JSON array, e.g., '[\"value1\",\"value2\"]')"
                        : "JSON array of strings",
                    defaultExpression: "null",
                    preProcessingCode: $"var {localVarName} = Sbroenne.ExcelMcp.Core.Utilities.ParameterTransforms.ParseJsonList({snakeName}, nameof({snakeName}));"));
            }
            else
            {
                // Direct passthrough — type matches between MCP and RouteAction
                var mcpType = ep.TypeName;

                // All exposed params are optional in MCP (not all actions use every param).
                // Use null as the category-wide omission sentinel, even when an individual action
                // has a Core default. Core dispatch applies that default after action validation.
                if (!mcpType.EndsWith("?"))
                    mcpType += "?";

                result.Add(new McpParameter(
                    name: snakeName,
                    mcpTypeName: mcpType,
                    routeActionParamName: ep.Name,
                    routeActionValue: snakeName,
                    description: description,
                    defaultExpression: "null",
                    preProcessingCode: null));
            }
        }

        return result;
    }

    private static string BuildParameterActionContext(ServiceInfo info, ExposedParameter parameter)
    {
        var actions = info.Methods
            .Where(method => method.Parameters.Any(methodParameter =>
                string.Equals(methodParameter.ExposedName ?? methodParameter.Name, parameter.Name, StringComparison.OrdinalIgnoreCase)))
            .Select(method => string.IsNullOrWhiteSpace(method.XmlDocSummary)
                ? method.ActionName
                : $"{method.ActionName}: {method.XmlDocSummary}")
            .Distinct(StringComparer.Ordinal);
        return string.Join(" ", actions);
    }

    private static string RenderMcpParameterNames(string description, List<ExposedParameter> parameters)
    {
        // Only known top-level camelCase inputs are translated. Nested JSON keys,
        // enum values, output fields, and ordinary words retain their own spelling.
        var names = parameters.Where(parameter => parameter.Name.Any(char.IsUpper))
            .ToDictionary(parameter => parameter.Name, parameter => StringHelper.ToSnakeCase(parameter.Name));
        var objectDepth = 0;
        return Regex.Replace(description, @"[{}]|\b[A-Za-z][A-Za-z0-9]*\b", match =>
        {
            if (match.Value == "{")
                objectDepth++;
            else if (match.Value == "}")
                objectDepth = Math.Max(0, objectDepth - 1);
            else if (objectDepth == 0 && names.TryGetValue(match.Value, out var name))
                return name;

            return match.Value;
        });
    }

    private static string BuildEnumAliasArguments(ParameterInfo parameter)
    {
        if (parameter.EnumAliases.Count == 0 || parameter.EnumTypeName == null)
            return string.Empty;

        return ", " + string.Join(
            ", ",
            parameter.EnumAliases.Select(alias =>
                $"(\"{EscapeStringLiteral(alias.Alias)}\", {parameter.EnumTypeName}.{alias.MemberName})"));
    }

    private static string GetClassName(ServiceInfo info)
    {
        return $"Excel{info.CategoryPascal}{(info.McpToolReadOnly ? "ReadOnly" : "")}Tool";
    }

    private static string GetMethodName(ServiceInfo info)
    {
        return $"Excel{info.CategoryPascal}{(info.McpToolReadOnly ? "ReadOnly" : "")}";
    }

    private static string GetActionTypeName(ServiceInfo info) =>
        $"Mcp{info.CategoryPascal}{(info.McpToolReadOnly ? "ReadOnly" : "")}Action";

    private static string[] WrapXmlDocLines(string text)
    {
        // Split on actual newlines, trim each line
        return text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.Trim())
            .Where(l => l.Length > 0)
            .ToArray();
    }

    private static string EscapeXml(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    /// <summary>
    /// Escapes a string for use inside a C# string literal (double-quoted).
    /// </summary>
    private static string EscapeStringLiteral(string text)
    {
        return text
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\r", "")
            .Replace("\n", " ");
    }

    /// <summary>
    /// Represents a parameter as it appears in the generated MCP tool method.
    /// </summary>
    private sealed class McpParameter
    {
        public string Name { get; }
        public string McpTypeName { get; }
        public string RouteActionParamName { get; }
        public string RouteActionValue { get; }
        public string? Description { get; }
        public string? DefaultExpression { get; }
        public string? PreProcessingCode { get; }

        public McpParameter(string name, string mcpTypeName, string routeActionParamName,
            string routeActionValue, string? description, string? defaultExpression,
            string? preProcessingCode)
        {
            Name = name;
            McpTypeName = mcpTypeName;
            RouteActionParamName = routeActionParamName;
            RouteActionValue = routeActionValue;
            Description = description;
            DefaultExpression = defaultExpression;
            PreProcessingCode = preProcessingCode;
        }
    }
}
