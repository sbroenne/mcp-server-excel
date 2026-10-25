using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace Sbroenne.ExcelMcp.Analyzers;

/// <summary>Checks repository-specific failure handling and COM safety.</summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class SafetyAnalyzer : DiagnosticAnalyzer
{
    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
        [Rules.EmptyCatch, Rules.BroadCatch, Rules.UnjustifiedDynamicCast,
            Rules.UnjustifiedSuppression, Rules.DirectDynamicCast, Rules.DynamicAccess];

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterCompilationStartAction(start =>
        {
            start.RegisterSyntaxNodeAction(AnalyzeDynamicType, SyntaxKind.IdentifierName);
            start.RegisterOperationAction(AnalyzeDynamicOperation,
                OperationKind.DynamicInvocation,
                OperationKind.DynamicMemberReference,
                OperationKind.DynamicIndexerAccess,
                OperationKind.DynamicObjectCreation);
            start.RegisterSyntaxNodeAction(AnalyzeInferredDynamicLocal,
                SyntaxKind.VariableDeclarator,
                SyntaxKind.ForEachStatement);
            start.RegisterSyntaxNodeAction(AnalyzeCast, SyntaxKind.CastExpression);
            if (start.Compilation.AssemblyName?.EndsWith(".Tests", StringComparison.Ordinal) == true) { return; }
            start.RegisterSyntaxNodeAction(AnalyzeCatch, SyntaxKind.CatchClause);
            start.RegisterSyntaxTreeAction(AnalyzePragmas);
            start.RegisterSymbolAction(AnalyzeSuppressions, SymbolKind.Method, SymbolKind.NamedType, SymbolKind.Field,
                SymbolKind.Property, SymbolKind.Event);
        });
    }

    private static void AnalyzeDynamicType(SyntaxNodeAnalysisContext context)
    {
        if (context.Node is not TypeSyntax type ||
            type.Parent is CastExpressionSyntax ||
            context.SemanticModel.GetSymbolInfo(type, context.CancellationToken).Symbol is not ITypeSymbol { TypeKind: TypeKind.Dynamic })
        {
            return;
        }
        context.ReportDiagnostic(Diagnostic.Create(Rules.DynamicAccess, type.GetLocation()));
    }

    private static void AnalyzeDynamicOperation(OperationAnalysisContext context) =>
        context.ReportDiagnostic(Diagnostic.Create(Rules.DynamicAccess, context.Operation.Syntax.GetLocation()));

    private static void AnalyzeInferredDynamicLocal(SyntaxNodeAnalysisContext context)
    {
        var inferred = context.Node switch
        {
            VariableDeclaratorSyntax variable when variable.Parent is VariableDeclarationSyntax
            { Type: IdentifierNameSyntax { Identifier.ValueText: "var" } } =>
                context.SemanticModel.GetDeclaredSymbol(variable, context.CancellationToken),
            ForEachStatementSyntax { Type: IdentifierNameSyntax { Identifier.ValueText: "var" } } forEach =>
                context.SemanticModel.GetDeclaredSymbol(forEach, context.CancellationToken),
            _ => null
        };
        if (inferred is ILocalSymbol { Type.TypeKind: TypeKind.Dynamic })
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.DynamicAccess, context.Node.GetLocation()));
        }
    }

    private static void AnalyzeCatch(SyntaxNodeAnalysisContext context)
    {
        var clause = (CatchClauseSyntax)context.Node;
        if (Propagates(clause.Block, context.SemanticModel)) { return; }
        if (clause.Filter is { } filter &&
            IsSpecificFilter(filter.FilterExpression, context.SemanticModel,
                clause.Declaration is { } declaration ? context.SemanticModel.GetDeclaredSymbol(declaration) : null,
                new HashSet<IMethodSymbol>(SymbolEqualityComparer.Default)))
        {
            return;
        }
        if (clause.Block.Statements.Count == 0)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.EmptyCatch, clause.CatchKeyword.GetLocation()));
            return;
        }
        var type = clause.Declaration is { } catchDeclaration
            ? context.SemanticModel.GetTypeInfo(catchDeclaration.Type).Type?.ToDisplayString()
            : "System.Exception";
        if (type is "System.Exception" or "System.SystemException" or "System.Runtime.InteropServices.COMException")
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.BroadCatch, clause.CatchKeyword.GetLocation(), type));
        }
    }

    private static bool Propagates(BlockSyntax block, SemanticModel model)
    {
        var flow = model.AnalyzeControlFlow(block);
        return flow is { Succeeded: true, EndPointIsReachable: false } &&
            flow.ExitPoints.Length == 0 && block.DescendantNodes(node =>
                node is not AnonymousFunctionExpressionSyntax and not LocalFunctionStatementSyntax)
                .Any(node => node is ThrowStatementSyntax or ThrowExpressionSyntax);
    }

    private static bool IsSpecificFilter(ExpressionSyntax expression, SemanticModel model, ISymbol? exception,
        HashSet<IMethodSymbol> visiting)
    {
        switch (expression)
        {
            case ParenthesizedExpressionSyntax parentheses:
                return IsSpecificFilter(parentheses.Expression, model, exception, visiting);
            case BinaryExpressionSyntax binary when binary.IsKind(SyntaxKind.LogicalOrExpression):
                return IsSpecificFilter(binary.Left, model, exception, visiting) && IsSpecificFilter(binary.Right, model, exception, visiting);
            case BinaryExpressionSyntax binary when binary.IsKind(SyntaxKind.LogicalAndExpression):
                return IsSpecificFilter(binary.Left, model, exception, visiting) ||
                    IsSpecificFilter(binary.Right, model, exception, visiting) ||
                    ExceptionAlias(binary.Left, model, exception) is { } alias &&
                    IsSpecificFilter(binary.Right, model, alias, visiting);
            case BinaryExpressionSyntax binary when binary.IsKind(SyntaxKind.EqualsExpression):
                return IsHResult(binary.Left, model, exception) && IsSpecificCode(binary.Right, model) ||
                    IsHResult(binary.Right, model, exception) && IsSpecificCode(binary.Left, model);
            case IsPatternExpressionSyntax pattern:
                return IsHResult(pattern.Expression, model, exception) && IsSpecificPattern(pattern.Pattern, model);
            case InvocationExpressionSyntax invocation:
                if (model.GetOperation(invocation) is not IInvocationOperation call ||
                    call.Arguments.FirstOrDefault(argument => RefersTo(argument.Value, exception))?.Parameter is not { } parameter ||
                    !visiting.Add(call.TargetMethod))
                {
                    return false;
                }
                var method = call.TargetMethod;
                try
                {
                    if (method.ContainingAssembly.Name == "Sbroenne.ExcelMcp.ComInterop" &&
                        method.ContainingType.ToDisplayString() == "Sbroenne.ExcelMcp.ComInterop.ComErrors" &&
                        method.Name == "IsMemberUnavailable" && method.ReturnType.SpecialType == SpecialType.System_Boolean)
                    {
                        return true;
                    }
                    foreach (var reference in method.DeclaringSyntaxReferences)
                    {
                        if (reference.GetSyntax() is not MethodDeclarationSyntax declaration) { continue; }
                        var body = declaration.ExpressionBody?.Expression ??
                            (declaration.Body?.Statements.Count == 1
                                ? (declaration.Body.Statements[0] as ReturnStatementSyntax)?.Expression
                                : null);
#pragma warning disable RS1030 // Inspect only source predicates actually called by a catch filter; the recursion guard bounds traversal.
                        if (body != null && IsSpecificFilter(body, model.Compilation.GetSemanticModel(body.SyntaxTree), parameter, visiting))
#pragma warning restore RS1030
                        {
                            return true;
                        }
                    }
                    return false;
                }
                finally { visiting.Remove(method); }
            default:
                return false;
        }
    }

    private static ISymbol? ExceptionAlias(ExpressionSyntax expression, SemanticModel model, ISymbol? exception)
    {
        while (expression is ParenthesizedExpressionSyntax parentheses) { expression = parentheses.Expression; }
        return expression is IsPatternExpressionSyntax
        {
            Expression: var source,
            Pattern: DeclarationPatternSyntax { Designation: SingleVariableDesignationSyntax variable }
        } && model.GetOperation(source) is { } operation && RefersTo(operation, exception)
            ? model.GetDeclaredSymbol(variable)
            : null;
    }

    private static bool IsHResult(ExpressionSyntax expression, SemanticModel model, ISymbol? exception) =>
        model.GetOperation(expression) is IPropertyReferenceOperation
        {
            Property: { Name: "HResult", ContainingType: { Name: "Exception", ContainingNamespace: { } ns } },
            Instance: { } instance
        } && ns.ToDisplayString() == "System" && RefersTo(instance, exception);

    private static bool RefersTo(IOperation value, ISymbol? symbol)
    {
        if (symbol == null) { return false; }
        while (value is IConversionOperation conversion) { value = conversion.Operand; }
        return value switch
        {
            ILocalReferenceOperation local => SymbolEqualityComparer.Default.Equals(local.Local, symbol),
            IParameterReferenceOperation parameter => SymbolEqualityComparer.Default.Equals(parameter.Parameter, symbol),
            _ => false
        };
    }

    private static bool IsSpecificCode(ExpressionSyntax expression, SemanticModel model) =>
        model.GetSymbolInfo(expression).Symbol is IFieldSymbol { IsConst: true, ConstantValue: int code } &&
        code < 0 && code != unchecked((int)0x800A03EC);

    private static bool IsSpecificPattern(PatternSyntax pattern, SemanticModel model) => pattern switch
    {
        ConstantPatternSyntax constant => IsSpecificCode(constant.Expression, model),
        BinaryPatternSyntax binary when binary.IsKind(SyntaxKind.OrPattern) =>
            IsSpecificPattern(binary.Left, model) && IsSpecificPattern(binary.Right, model),
        ParenthesizedPatternSyntax parenthesized => IsSpecificPattern(parenthesized.Pattern, model),
        _ => false
    };

    private static void AnalyzeCast(SyntaxNodeAnalysisContext context)
    {
        var cast = (CastExpressionSyntax)context.Node;
        var target = context.SemanticModel.GetTypeInfo(cast.Type).Type;
        var source = context.SemanticModel.GetTypeInfo(cast.Expression).Type;
        if (target?.TypeKind == TypeKind.Dynamic)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.UnjustifiedDynamicCast, cast.Type.GetLocation()));
        }
        if (source?.TypeKind != TypeKind.Dynamic || target == null) { return; }
        var conversion = target.TypeKind == TypeKind.Enum ? "Int32" : target.SpecialType switch
        {
            SpecialType.System_Int32 => "Int32",
            SpecialType.System_Int64 => "Int64",
            SpecialType.System_Int16 => "Int16",
            SpecialType.System_Byte => "Byte",
            SpecialType.System_SByte => "SByte",
            SpecialType.System_UInt16 => "UInt16",
            SpecialType.System_UInt32 => "UInt32",
            SpecialType.System_UInt64 => "UInt64",
            SpecialType.System_Single => "Single",
            SpecialType.System_Double => "Double",
            SpecialType.System_Decimal => "Decimal",
            _ => null
        };
        if (conversion != null)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.DirectDynamicCast, cast.Type.GetLocation(),
                conversion, target.ToDisplayString()));
        }
    }

    private static void AnalyzePragmas(SyntaxTreeAnalysisContext context)
    {
        var root = context.Tree.GetRoot(context.CancellationToken);
        foreach (var pragma in root.DescendantTrivia(descendIntoTrivia: true)
            .Select(trivia => trivia.GetStructure()).OfType<PragmaWarningDirectiveTriviaSyntax>())
        {
            if (!pragma.DisableOrRestoreKeyword.IsKind(SyntaxKind.DisableKeyword)) { continue; }
            var line = context.Tree.GetText(context.CancellationToken).Lines.GetLineFromPosition(pragma.SpanStart).ToString();
            var comment = line.IndexOf("//", StringComparison.Ordinal);
            if (comment >= 0 && !string.IsNullOrWhiteSpace(line.Substring(comment + 2))) { continue; }
            context.ReportDiagnostic(Diagnostic.Create(Rules.UnjustifiedSuppression, pragma.GetLocation(),
                "#pragma warning disable", "a trailing // explanation"));
        }
    }

    private static void AnalyzeSuppressions(SymbolAnalysisContext context)
    {
        foreach (var attribute in context.Symbol.GetAttributes())
        {
            if (attribute.AttributeClass?.ToDisplayString() != "System.Diagnostics.CodeAnalysis.SuppressMessageAttribute")
            {
                continue;
            }
            if (attribute.NamedArguments.Any(pair => pair.Key == "Justification" &&
                pair.Value.Value is string reason && !string.IsNullOrWhiteSpace(reason)))
            {
                continue;
            }
            var location = attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation();
            context.ReportDiagnostic(Diagnostic.Create(Rules.UnjustifiedSuppression, location,
                "SuppressMessage", "Justification = \"...\""));
        }
    }
}
