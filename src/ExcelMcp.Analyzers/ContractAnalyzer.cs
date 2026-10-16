using System;
using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

namespace Sbroenne.ExcelMcp.Analyzers;

/// <summary>Checks source contracts and the generated MCP method shape.</summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class ContractAnalyzer : DiagnosticAnalyzer
{
    private static readonly DiagnosticDescriptor SynchronousCore = new(
        "EXL013", "Core service contracts run synchronously",
        "'{0}' returns an awaitable. Core commands execute synchronously on the Excel STA; keep async transport in Service and adapters.",
        "Contracts", DiagnosticSeverity.Error, isEnabledByDefault: true);
    private static readonly DiagnosticDescriptor McpMethod = new(
        "EXL014", "MCP tools preserve SDK results and cancellation",
        "MCP tool '{0}' must return Task<CallToolResult> from the SDK and accept a CancellationToken",
        "Contracts", DiagnosticSeverity.Error, isEnabledByDefault: true);
    private static readonly DiagnosticDescriptor CamelCase = new(
        "EXL016", "Service contract parameters use camelCase",
        "Parameter '{0}' must use camelCase because CLI, Service and MCP names derive from this contract",
        "Contracts", DiagnosticSeverity.Error, isEnabledByDefault: true);

    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [SynchronousCore, McpMethod, CamelCase];

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze | GeneratedCodeAnalysisFlags.ReportDiagnostics);
        context.EnableConcurrentExecution();
        context.RegisterSymbolAction(AnalyzeMethod, SymbolKind.Method);
    }

    private static void AnalyzeMethod(SymbolAnalysisContext context)
    {
        if (context.Compilation.AssemblyName?.EndsWith(".Tests", StringComparison.Ordinal) == true) { return; }
        var method = (IMethodSymbol)context.Symbol;
        if (method.ContainingType.TypeKind == TypeKind.Interface &&
            method.ContainingType.GetAttributes().Any(a =>
                a.AttributeClass?.ToDisplayString() == "Sbroenne.ExcelMcp.Core.Attributes.ServiceCategoryAttribute"))
        {
            if (method.ReturnType.ContainingNamespace?.ToDisplayString() == "System.Threading.Tasks" &&
                method.ReturnType.Name is "Task" or "ValueTask")
            {
                context.ReportDiagnostic(Diagnostic.Create(SynchronousCore, method.Locations.FirstOrDefault(), method.Name));
            }
            foreach (var parameter in method.Parameters.Where(parameter => parameter.Name.Length == 0 ||
                !char.IsLower(parameter.Name[0]) || parameter.Name.Skip(1).Any(character => !char.IsLetterOrDigit(character))))
            {
                context.ReportDiagnostic(Diagnostic.Create(CamelCase, parameter.Locations.FirstOrDefault(), parameter.Name));
            }
        }
        if (!method.GetAttributes().Any(a =>
            a.AttributeClass?.ToDisplayString() == "ModelContextProtocol.Server.McpServerToolAttribute"))
        {
            return;
        }
        if (method.ReturnType is not INamedTypeSymbol { Name: "Task", Arity: 1 } task ||
            task.ContainingNamespace.ToDisplayString() != "System.Threading.Tasks" ||
            task.TypeArguments[0].ToDisplayString() != "ModelContextProtocol.Protocol.CallToolResult" ||
            !method.Parameters.Any(parameter => parameter.Type.ToDisplayString() == "System.Threading.CancellationToken"))
        {
            context.ReportDiagnostic(Diagnostic.Create(McpMethod, method.Locations.FirstOrDefault(), method.Name));
        }
    }
}
