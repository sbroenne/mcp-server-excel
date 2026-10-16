using System;
using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace Sbroenne.ExcelMcp.Analyzers;

/// <summary>Checks acquired COM intermediates using symbol types rather than property-name regexes.</summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class ComAccessAnalyzer : DiagnosticAnalyzer
{
    private static readonly ImmutableHashSet<string> DynamicComProperties = ImmutableHashSet.Create(
        "Rows", "Columns", "ListColumns", "ListRows", "Range", "TableRange1", "TableRange2",
        "ChartArea", "Parent", "Application", "Cells");
    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
        [Rules.ChainedComAccess, Rules.DiscardedComResult, Rules.WorkbookPackageAccess];

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterCompilationStartAction(start =>
        {
            if (start.Compilation.AssemblyName?.EndsWith(".Tests", StringComparison.Ordinal) == true) { return; }
            start.RegisterOperationAction(AnalyzeProperty, OperationKind.PropertyReference);
            start.RegisterOperationAction(AnalyzeInvocation, OperationKind.Invocation);
            start.RegisterOperationAction(AnalyzeObjectCreation, OperationKind.ObjectCreation);
            start.RegisterSyntaxNodeAction(AnalyzeLiteral, SyntaxKind.StringLiteralExpression);
            start.RegisterSyntaxNodeAction(AnalyzeDynamicChain, SyntaxKind.SimpleMemberAccessExpression);
        });
    }

    private static void AnalyzeDynamicChain(SyntaxNodeAnalysisContext context)
    {
        var member = (MemberAccessExpressionSyntax)context.Node;
        if (!DynamicComProperties.Contains(member.Name.Identifier.ValueText) ||
            context.SemanticModel.GetTypeInfo(member.Expression).Type?.TypeKind != TypeKind.Dynamic)
        {
            return;
        }
        SyntaxNode node = member;
        while (node.Parent is ParenthesizedExpressionSyntax or ElementAccessExpressionSyntax) { node = node.Parent; }
        if (node.Parent is MemberAccessExpressionSyntax outer && outer.Expression == node)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.ChainedComAccess, member.GetLocation(), member.ToString()));
        }
    }

    private static bool IsExcel(ITypeSymbol? type) =>
        type?.TypeKind is TypeKind.Interface or TypeKind.Class &&
        type.ContainingNamespace?.ToDisplayString() == "Microsoft.Office.Interop.Excel";

    private static void AnalyzeProperty(OperationAnalysisContext context)
    {
        var property = (IPropertyReferenceOperation)context.Operation;
        if (!IsExcel(property.Type) || property.Property.Name is "App" or "Book" ||
            property.Parent is not { } parent)
        {
            return;
        }
        while (parent is IConversionOperation or IParenthesizedOperation)
        {
            parent = parent.Parent;
        }
        if (parent is IPropertyReferenceOperation or IInvocationOperation or IArrayElementReferenceOperation)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.ChainedComAccess, property.Syntax.GetLocation(),
                property.Syntax.ToString()));
        }
    }

    private static void AnalyzeInvocation(OperationAnalysisContext context)
    {
        var invocation = (IInvocationOperation)context.Operation;
        if (IsArchiveApi(invocation.TargetMethod.ContainingType))
        {
            ReportArchiveAccess(context, invocation.Syntax, invocation.TargetMethod.ToDisplayString());
            return;
        }
        if (!IsExcel(invocation.Type)) { return; }
        IOperation? parent = invocation.Parent;
        while (parent is IConversionOperation or IParenthesizedOperation) { parent = parent.Parent; }
        if (parent is IExpressionStatementOperation)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.DiscardedComResult, invocation.Syntax.GetLocation(),
                invocation.TargetMethod.Name));
        }
        else if (parent is IPropertyReferenceOperation or IInvocationOperation)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.ChainedComAccess, invocation.Syntax.GetLocation(),
                invocation.Syntax.ToString()));
        }
    }

    private static void AnalyzeObjectCreation(OperationAnalysisContext context)
    {
        var creation = (IObjectCreationOperation)context.Operation;
        if (IsArchiveApi(creation.Type))
        {
            ReportArchiveAccess(context, creation.Syntax, creation.Type!.ToDisplayString());
        }
    }

    private static bool IsArchiveApi(ITypeSymbol? type) =>
        type is not null &&
        type.ContainingNamespace?.ToDisplayString() == "System.IO.Compression" &&
        type.Name is "ZipArchive" or "ZipFile";

    private static void ReportArchiveAccess(OperationAnalysisContext context, SyntaxNode syntax, string symbol)
    {
        string path = syntax.SyntaxTree.FilePath.Replace('\\', '/');
        if (path.EndsWith(
                "src/ExcelMcp.Core/Commands/Settings/NativeQueryApprovalSettings.cs",
                StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        context.ReportDiagnostic(Diagnostic.Create(Rules.WorkbookPackageAccess, syntax.GetLocation(), symbol));
    }

    private static void AnalyzeLiteral(SyntaxNodeAnalysisContext context)
    {
        var literal = (LiteralExpressionSyntax)context.Node;
        var value = literal.Token.ValueText.Replace('\\', '/');
        var workbookPart = value.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
            value.Split('/').Any(segment => segment.Equals("xl", StringComparison.OrdinalIgnoreCase)) &&
            value.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
        if (workbookPart)
        {
            context.ReportDiagnostic(Diagnostic.Create(Rules.WorkbookPackageAccess, literal.GetLocation(), literal.Token.ValueText));
        }
    }
}
