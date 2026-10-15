using System.Collections.Concurrent;
using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;
using Microsoft.CodeAnalysis.Text;
using System.Xml;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Analyzers;

[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class ExcelSafetyAnalyzer : DiagnosticAnalyzer
{
    public const string DynamicId = "EXCEL001";
    public const string NumericId = "EXCEL002";
    public const string ComId = "EXCEL003";
    public const string SuccessId = "EXCEL004";
    public const string PackageId = "EXCEL006";
    public const string CompletedAreaKey = "excelmcp_typed_excel";
    private static readonly string[] WorkbookExtensions = { ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm" };

    private static readonly DiagnosticDescriptor DynamicRule = Rule(DynamicId, "Use typed Excel access",
        "Use typed Excel PIA access; each verified runtime gap needs an adjacent '// PIA gap:' comment with member/runtime evidence");
    private static readonly DiagnosticDescriptor NumericRule = Rule(NumericId, "Convert dynamic numeric values",
        "Use Convert.* for numeric or enum values obtained through dynamic Excel access");
    private static readonly DiagnosticDescriptor ComRule = Rule(ComId, "Capture acquired Excel COM references",
        "Capture the acquired Excel COM reference in a local and release it in finally; do not chain or discard it");
    private static readonly DiagnosticDescriptor SuccessRule = Rule(SuccessId, "Successful results cannot contain errors",
        "This ResultBase is returned with Success=true and a nonempty ErrorMessage; clear the error or report failure");
    private static readonly DiagnosticDescriptor PackageRule = Rule(PackageId, "Use Excel COM for workbook contents",
        "Use Excel COM in production; workbook package part paths and XML read from ZIP entries belong only in tests");

    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
        ImmutableArray.Create(DynamicRule, NumericRule, ComRule, SuccessRule, PackageRule);

    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterSyntaxNodeAction(CheckDynamic, SyntaxKind.IdentifierName);
        context.RegisterOperationAction(CheckConversion, OperationKind.Conversion);
        context.RegisterOperationAction(CheckCom, OperationKind.PropertyReference, OperationKind.Invocation, OperationKind.Conversion);
        context.RegisterSyntaxNodeAction(CheckLiteral, SyntaxKind.StringLiteralExpression);
        context.RegisterCompilationStartAction(RegisterZipXml);
        context.RegisterCompilationAction(CheckBuildLiterals);
        context.RegisterSyntaxNodeAction(CheckInitializer, SyntaxKind.ObjectCreationExpression, SyntaxKind.ImplicitObjectCreationExpression);
        context.RegisterSyntaxNodeAction(CheckReturn, SyntaxKind.ReturnStatement);
    }

    private static DiagnosticDescriptor Rule(string id, string title, string message) =>
        new(id, title, message, "ExcelSafety", DiagnosticSeverity.Error, isEnabledByDefault: true);

    private static bool Completed(SyntaxTree tree, AnalyzerConfigOptionsProvider options) =>
        options.GetOptions(tree).TryGetValue(CompletedAreaKey, out var value) &&
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    private static void CheckDynamic(SyntaxNodeAnalysisContext context)
    {
        var node = (IdentifierNameSyntax)context.Node;
        if (node.Identifier.ValueText != "dynamic" ||
            context.SemanticModel.GetTypeInfo(node, context.CancellationToken).Type?.TypeKind != TypeKind.Dynamic)
        {
            return;
        }
        var symbol = context.SemanticModel.GetSymbolInfo(node, context.CancellationToken).Symbol;
        if (symbol != null && symbol is not ITypeSymbol) { return; }

        var completed = Completed(node.SyntaxTree, context.Options.AnalyzerConfigOptionsProvider);
        if (!completed && !(node.Parent is CastExpressionSyntax cast && cast.Parent is ParenthesizedExpressionSyntax))
        {
            return;
        }

        if (!HasGapComment(node, completed, context.SemanticModel))
        {
            context.ReportDiagnostic(Diagnostic.Create(DynamicRule, node.GetLocation()));
        }
    }

    private static bool HasGapComment(SyntaxNode node, bool completed, SemanticModel model)
    {
        // A comment belongs to the declaration/statement containing this occurrence,
        // not a prior unrelated statement or a blanket file exemption.
        var owner = node.AncestorsAndSelf().FirstOrDefault(n =>
            n is StatementSyntax || n is MemberDeclarationSyntax) ?? node;
        var boundaryTrivia = owner.Parent is ElseClauseSyntax clause ? clause.GetLeadingTrivia() : default;
        var comments = owner.GetLeadingTrivia().Concat(node.GetLeadingTrivia()).Concat(boundaryTrivia)
            .Where(t => t.IsKind(SyntaxKind.SingleLineCommentTrivia));
        foreach (var trivia in comments)
        {
            var comment = trivia.ToString();
            const string prefix = "// PIA gap:";
            if (comment.StartsWith(prefix, StringComparison.Ordinal) &&
                comment.Substring(prefix.Length).Trim().Length > 0 &&
                (!completed || ReviewedGap(comment) && GapMatchesTypedReceiver(node, comment, model)))
            {
                return true;
            }
            if (!completed && (comment.StartsWith("// TODO:", StringComparison.Ordinal) ||
                comment.StartsWith("// Reason:", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
        }
        return false;
    }

    private static bool ReviewedGap(string comment)
    {
        // These are the verified runtime/dependency gaps, not generic migration
        // reasons. Keep member-level evidence in docs/PIA-COVERAGE.md.
        var evidence = new[]
        {
            "Office.Core", "Office-core", "office.dll", "ADODB", "Microsoft.Vbe.Interop", "VBE project",
            "VBA editor object model", "PivotCache.ADOConnection",
            "Application.Run", "TextConnection.BackgroundQuery", "TextConnection.RefreshOnFileOpen",
            "TextConnection.SavePassword", "TextConnection.RefreshPeriod", "TextConnection.CommandText",
            "WorkbookConnection.WebConnection", "WorkbookConnection has no WebConnection",
            "OLAP RefreshDate", "ModelTableColumn.IsCalculatedColumn", "typed SourceName is a string",
            "PivotField.VisibleItems() fails without an index",
            "typed Open signature injects an LCID", "typed Save signature injects an LCID"
        };
        return evidence.Any(item => comment.IndexOf(item, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static bool GapMatchesTypedReceiver(SyntaxNode node, string evidence, SemanticModel model)
    {
        var cast = node.Parent as CastExpressionSyntax;
        ExpressionSyntax? operand = cast?.Expression ??
            node.Ancestors().OfType<VariableDeclarationSyntax>().FirstOrDefault()?.Variables.FirstOrDefault()?.Initializer?.Value;
        if (operand == null) { return true; }
        while (operand is ParenthesizedExpressionSyntax || operand is CastExpressionSyntax)
        {
            operand = operand is ParenthesizedExpressionSyntax p ? p.Expression : ((CastExpressionSyntax)operand).Expression;
        }
        var type = model.GetTypeInfo(operand).Type;
        if (!IsExcel(type)) { return true; } // External IDispatch declarations retain explicit reviewed evidence.
        var receiver = type!.Name;
        var member = cast?.Ancestors().OfType<MemberAccessExpressionSyntax>().FirstOrDefault()?.Name.Identifier.ValueText;
        return receiver switch
        {
            "Shape" => member is "Type" or "Connector" or "AutoShapeType" or "Fill" or "Line" or "ConnectorFormat" or "Visible" or "LockAspectRatio" ||
                member == null && evidence.IndexOf("Office", StringComparison.OrdinalIgnoreCase) >= 0,
            "Shapes" or "ShapeRange" => member is "AddPicture" or "AddShape" or "AddTextbox" or "AddConnector" or "Align" or "Distribute" or "ZOrder" ||
                member == null && evidence.IndexOf("Office", StringComparison.OrdinalIgnoreCase) >= 0,
            "ChartFormat" => member is "Fill" or "Line",
            "TextConnection" => member is "BackgroundQuery" or "RefreshOnFileOpen" or "SavePassword" or "RefreshPeriod" or "CommandText",
            "WorkbookConnection" => member is "WebConnection" or "ConnectionString",
            "OLEDBConnection" => member == "RefreshDate",
            "ModelTableColumn" => member == "IsCalculatedColumn",
            "PivotField" => member is "SourceName" or "VisibleItems",
            "Application" or "_Application" => member is "Run" or "VBE" or "AutomationSecurity" || member == null,
            "Workbook" or "_Workbook" => member is "VBProject" or "Theme" or "DocumentInspectors" or "Save",
            "Workbooks" => evidence.IndexOf("LCID", StringComparison.OrdinalIgnoreCase) >= 0,
            _ => false
        };
    }

    private static void CheckConversion(OperationAnalysisContext context)
    {
        var conversion = (IConversionOperation)context.Operation;
        if (conversion.IsImplicit || conversion.Operand.Type?.TypeKind != TypeKind.Dynamic ||
            !Completed(conversion.Syntax.SyntaxTree, context.Options.AnalyzerConfigOptionsProvider))
        {
            return;
        }
        var type = conversion.Type;
        if (type is INamedTypeSymbol nullable && nullable.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            type = nullable.TypeArguments[0];
        }
        if (type?.TypeKind == TypeKind.Enum || type?.SpecialType is
            SpecialType.System_Byte or SpecialType.System_SByte or SpecialType.System_Int16 or
            SpecialType.System_UInt16 or SpecialType.System_Int32 or SpecialType.System_UInt32 or
            SpecialType.System_Int64 or SpecialType.System_UInt64 or SpecialType.System_Single or
            SpecialType.System_Double or SpecialType.System_Decimal or SpecialType.System_IntPtr or SpecialType.System_UIntPtr)
        {
            context.ReportDiagnostic(Diagnostic.Create(NumericRule, conversion.Syntax.GetLocation()));
        }
    }

    private static bool IsExcel(ITypeSymbol? type) =>
        type?.IsReferenceType == true &&
        type.ContainingNamespace?.ToDisplayString() == "Microsoft.Office.Interop.Excel";

    private static bool Borrowed(IOperation operation) =>
        operation is IPropertyReferenceOperation property &&
        property.Property.ContainingType.ToDisplayString() == "Sbroenne.ExcelMcp.ComInterop.Session.ExcelContext" &&
        property.Property.Name is "App" or "Book" ||
        operation is IInvocationOperation invocation && invocation.TargetMethod.Name == "GetWorkbook" &&
        invocation.TargetMethod.ContainingType.ToDisplayString() is
            "Sbroenne.ExcelMcp.ComInterop.Session.ExcelBatch" or "Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch";

    private static void CheckCom(OperationAnalysisContext context)
    {
        var operation = context.Operation;
        if (!IsExcel(operation.Type) || Borrowed(operation) ||
            operation is IPropertyReferenceOperation managedProperty && !IsExcel(managedProperty.Property.ContainingType) ||
            operation is IInvocationOperation managedMethod && !IsExcel(managedMethod.TargetMethod.ContainingType))
        {
            return;
        }
        if (operation is IConversionOperation conversion)
        {
            IOperation acquisition = conversion;
            while (acquisition is IConversionOperation cast && !cast.Conversion.IsUserDefined &&
                (cast.Conversion.IsIdentity || cast.Conversion.IsReference))
            {
                acquisition = cast.Operand;
                // A typed acquisition is checked at its own operation, not again at enclosing casts.
                if (IsExcel(acquisition.Type)) { return; }
            }
            if (!(acquisition is IPropertyReferenceOperation propertyAcquisition && IsExcel(propertyAcquisition.Property.ContainingType) ||
                acquisition is IInvocationOperation methodAcquisition && IsExcel(methodAcquisition.TargetMethod.ContainingType)))
            {
                return;
            }
        }
        var parent = operation.Parent;
        while (parent is IConversionOperation || parent is IParenthesizedOperation)
        {
            parent = parent.Parent;
        }
        var chained = parent is IPropertyReferenceOperation property && property.Instance != null ||
            parent is IInvocationOperation invocation && invocation.Instance != null ||
            parent is IDynamicMemberReferenceOperation;
        var discarded = parent is IExpressionStatementOperation ||
            parent is ISimpleAssignmentOperation assignment && assignment.Target is IDiscardOperation;
        if (chained || discarded)
        {
            context.ReportDiagnostic(Diagnostic.Create(ComRule, operation.Syntax.GetLocation()));
        }
    }

    private static bool WorkbookPart(string text)
    {
        text = text.Replace('\\', '/').TrimStart('/');
        return string.Equals(text, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
            text.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) &&
            text.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
    }

    private static void CheckLiteral(SyntaxNodeAnalysisContext context)
    {
        var literal = (LiteralExpressionSyntax)context.Node;
        if (WorkbookPart(literal.Token.ValueText))
        {
            context.ReportDiagnostic(Diagnostic.Create(PackageRule, literal.GetLocation()));
        }
    }

    private static void CheckBuildLiterals(CompilationAnalysisContext context)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in context.Options.AdditionalFiles)
        {
            if (!visited.Add(file.Path)) { continue; }
            if (!file.Path.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase) &&
                !file.Path.EndsWith(".props", StringComparison.OrdinalIgnoreCase) &&
                !file.Path.EndsWith(".targets", StringComparison.OrdinalIgnoreCase)) { continue; }
            var text = file.GetText(context.CancellationToken);
            if (text == null) { continue; }
            XDocument document;
            try { document = XDocument.Parse(text.ToString(), LoadOptions.SetLineInfo); }
            catch (XmlException) { continue; } // MSBuild reports malformed project XML.
            foreach (var element in document.Descendants())
            {
                var candidates = element.Attributes().Select(a => (Value: a.Value, Info: (IXmlLineInfo)a))
                    .Concat(element.HasElements ? Enumerable.Empty<(string Value, IXmlLineInfo Info)>() :
                        new[] { (element.Value, (IXmlLineInfo)element) });
                foreach (var candidate in candidates.Where(c => WorkbookPart(c.Value)))
                {
                    var line = Math.Max(candidate.Info.LineNumber - 1, 0);
                    var position = new LinePosition(line, Math.Max(candidate.Info.LinePosition - 1, 0));
                    var location = Location.Create(file.Path, new TextSpan(text.Lines[line].Start + position.Character, 0),
                        new LinePositionSpan(position, position));
                    context.ReportDiagnostic(Diagnostic.Create(PackageRule, location));
                }
            }
        }
    }

    private static void RegisterZipXml(CompilationStartAnalysisContext context)
    {
        var values = new ConcurrentDictionary<ISymbol, ImmutableArray<IOperation>>(SymbolEqualityComparer.Default);
        var sinks = new ConcurrentBag<IInvocationOperation>();
        var calls = new ConcurrentBag<IInvocationOperation>();
        context.RegisterOperationBlockAction(block =>
        {
            var operations = block.OperationBlocks.SelectMany(OperationExtensions.DescendantsAndSelf).ToArray();
            values[block.OwningSymbol] = operations.OfType<IReturnOperation>()
                .Where(r => r.ReturnedValue != null).Select(r => r.ReturnedValue!).ToImmutableArray();
            foreach (var variable in operations.OfType<IVariableDeclaratorOperation>())
            {
                if (variable.Initializer != null)
                {
                    values[variable.Symbol] = ImmutableArray.Create(variable.Initializer.Value);
                }
            }
            foreach (var invocation in operations.OfType<IInvocationOperation>())
            {
                calls.Add(invocation);
                if (invocation.TargetMethod.ContainingNamespace.ToDisplayString().StartsWith("System.Xml", StringComparison.Ordinal) &&
                    invocation.TargetMethod.Name is "Load" or "Create" or "Parse" or "LoadXml")
                {
                    sinks.Add(invocation);
                }
            }
        });
        context.RegisterCompilationEndAction(end =>
        {
            foreach (var invocation in calls)
            {
                foreach (var argument in invocation.Arguments)
                {
                    if (argument.Parameter != null && invocation.TargetMethod.DeclaringSyntaxReferences.Length > 0)
                    {
                        values.AddOrUpdate(argument.Parameter.OriginalDefinition, ImmutableArray.Create(argument.Value),
                            (_, existing) => existing.Add(argument.Value));
                    }
                }
            }
            foreach (var invocation in sinks)
            {
                if (invocation.Arguments.Any(a => FromZip(a.Value, values, new HashSet<ISymbol>(SymbolEqualityComparer.Default), 0)))
                {
                    end.ReportDiagnostic(Diagnostic.Create(PackageRule, invocation.Syntax.GetLocation()));
                }
            }
        });
    }

    private static bool FromZip(IOperation? operation,
        ConcurrentDictionary<ISymbol, ImmutableArray<IOperation>> values, HashSet<ISymbol> visited, int depth)
    {
        if (operation == null || depth > 16) { return false; }
        if (operation.ConstantValue is { HasValue: true, Value: string text } && WorkbookPart(text))
        {
            return true;
        }
        switch (operation)
        {
            case IConversionOperation conversion:
                return FromZip(conversion.Operand, values, visited, depth + 1);
            case IObjectCreationOperation creation:
                if (creation.Type?.ToDisplayString() == "System.IO.Compression.ZipArchive" &&
                    creation.Arguments.Any(a => WorkbookPath(a.Value, values,
                        new HashSet<ISymbol>(SymbolEqualityComparer.Default), 0)))
                {
                    return true;
                }
                return creation.Arguments.Any(a => FromZip(a.Value, values, visited, depth + 1));
            case IPropertyReferenceOperation property:
                return FromZip(property.Instance, values, visited, depth + 1);
            case IInvocationOperation invocation:
                if (invocation.TargetMethod.ContainingType.ToDisplayString() == "System.IO.Compression.ZipFile" &&
                    invocation.TargetMethod.Name is "Open" or "OpenRead" &&
                    invocation.Arguments.Any(a => WorkbookPath(a.Value, values,
                        new HashSet<ISymbol>(SymbolEqualityComparer.Default), 0)))
                {
                    return true;
                }
                if (FromZip(invocation.Instance, values, visited, depth + 1) ||
                    invocation.Arguments.Any(a => FromZip(a.Value, values, visited, depth + 1)))
                {
                    return true;
                }
                if (!visited.Add(invocation.TargetMethod)) { return false; }
                return values.TryGetValue(invocation.TargetMethod, out var returns) &&
                    returns.Any(e => FromZip(e, values, visited, depth + 1));
            case ILocalReferenceOperation local:
                if (!visited.Add(local.Local)) { return false; }
                return values.TryGetValue(local.Local, out var initializers) &&
                    initializers.Any(e => FromZip(e, values, visited, depth + 1));
            case IParameterReferenceOperation parameter:
                if (!visited.Add(parameter.Parameter.OriginalDefinition)) { return false; }
                return values.TryGetValue(parameter.Parameter.OriginalDefinition, out var arguments) &&
                    arguments.Any(e => FromZip(e, values, visited, depth + 1));
            default:
                return false;
        }
    }

    private static bool WorkbookPath(IOperation operation,
        ConcurrentDictionary<ISymbol, ImmutableArray<IOperation>> values, HashSet<ISymbol> visited, int depth)
    {
        if (depth > 16) { return false; }
        if (operation.ConstantValue is { HasValue: true, Value: string path })
        {
            return WorkbookExtensions.Any(extension => path.EndsWith(extension, StringComparison.OrdinalIgnoreCase));
        }
        ISymbol? symbol = operation switch
        {
            ILocalReferenceOperation local => local.Local,
            IParameterReferenceOperation parameter => parameter.Parameter.OriginalDefinition,
            IInvocationOperation invocation => invocation.TargetMethod,
            _ => null
        };
        if (operation is IConversionOperation conversion)
        {
            return WorkbookPath(conversion.Operand, values, visited, depth + 1);
        }
        if (operation is IInvocationOperation file &&
            file.TargetMethod.ContainingType.ToDisplayString() == "System.IO.File" &&
            file.TargetMethod.Name is "Open" or "OpenRead")
        {
            return file.Arguments.Any(argument => WorkbookPath(argument.Value, values, visited, depth + 1));
        }
        if (operation is IObjectCreationOperation creation && creation.Type?.ToDisplayString() == "System.IO.FileStream")
        {
            return creation.Arguments.Any(argument => WorkbookPath(argument.Value, values, visited, depth + 1));
        }
        return symbol != null && visited.Add(symbol) && values.TryGetValue(symbol, out var sources) &&
            sources.Any(source => WorkbookPath(source, values, visited, depth + 1));
    }

    private static bool IsResult(ITypeSymbol? type)
    {
        for (var current = type as INamedTypeSymbol; current != null; current = current.BaseType)
        {
            if (current.ToDisplayString() == "Sbroenne.ExcelMcp.Core.Models.ResultBase") { return true; }
        }
        return false;
    }

    private static void CheckInitializer(SyntaxNodeAnalysisContext context)
    {
        if (!IsResult(context.SemanticModel.GetTypeInfo(context.Node).Type)) { return; }
        // Local construction is provisional; only the terminal returned state
        // establishes a conflict. Exposed/returned creations have no repair window.
        SyntaxNode binding = context.Node;
        while (binding.Parent is ParenthesizedExpressionSyntax or CastExpressionSyntax) { binding = binding.Parent; }
        if (binding.Parent is EqualsValueClauseSyntax { Parent: VariableDeclaratorSyntax variable } &&
            context.SemanticModel.GetDeclaredSymbol(variable) is ILocalSymbol ||
            binding.Parent is AssignmentExpressionSyntax localAssignment &&
            localAssignment.IsKind(SyntaxKind.SimpleAssignmentExpression) &&
            localAssignment.Parent is ExpressionStatementSyntax &&
            ReferencedLocal(localAssignment.Left, context.SemanticModel) != null)
        {
            return;
        }
        var initializer = (context.Node as BaseObjectCreationExpressionSyntax)?.Initializer;
        if (initializer == null) { return; }
        bool? success = null;
        bool? error = null;
        foreach (var assignment in initializer.Expressions.OfType<AssignmentExpressionSyntax>())
        {
            SetState(assignment.Left.ToString(), assignment.Right, context.SemanticModel, ref success, ref error);
        }
        if (success == true && error == true)
        {
            context.ReportDiagnostic(Diagnostic.Create(SuccessRule, context.Node.GetLocation()));
        }
    }

    private static void SetState(string property, ExpressionSyntax value, SemanticModel model, ref bool? success, ref bool? error)
    {
        var constant = model.GetConstantValue(value);
        if (property == "Success") { success = constant.HasValue && constant.Value is bool flag ? flag : null; }
        if (property == "ErrorMessage")
        {
            error = constant.HasValue ? constant.Value is string text && text.Length > 0 :
                value.ToString() == "string.Empty" ? false : (bool?)null;
        }
    }

    private static void CheckReturn(SyntaxNodeAnalysisContext context)
    {
        var returned = (ReturnStatementSyntax)context.Node;
        if (returned.Expression == null ||
            ReferencedLocal(returned.Expression, context.SemanticModel) is not ILocalSymbol local)
        {
            return;
        }
        var block = returned.Ancestors()
            .TakeWhile(node => node is not BaseMethodDeclarationSyntax and not LocalFunctionStatementSyntax and
                not AnonymousFunctionExpressionSyntax and not AccessorDeclarationSyntax)
            .OfType<BlockSyntax>().LastOrDefault();
        if (block == null) { return; }
        var states = new Dictionary<ILocalSymbol, ResultState>(SymbolEqualityComparer.Default);
        Dictionary<ILocalSymbol, ResultState>? terminal = null;
        AnalyzeStatements(block, context.SemanticModel, states, returned, ref terminal, 0);
        if (terminal != null && terminal.TryGetValue(local, out var returnedState) && !returnedState.Escaped &&
            returnedState.Success == true && returnedState.Error == true)
        {
            context.ReportDiagnostic(Diagnostic.Create(SuccessRule, returned.GetLocation()));
        }
    }

    private static Dictionary<ILocalSymbol, ResultState>? AnalyzeStatements(BlockSyntax block, SemanticModel model,
        Dictionary<ILocalSymbol, ResultState> states, ReturnStatementSyntax target,
        ref Dictionary<ILocalSymbol, ResultState>? terminal, int depth)
    {
        foreach (var statement in block.Statements)
        {
            var continuation = AnalyzeStatement(statement, model, states, target, ref terminal, depth);
            if (continuation == null) { return null; }
            states = continuation;
        }
        return states;
    }

    private static Dictionary<ILocalSymbol, ResultState>? AnalyzeStatement(StatementSyntax statement, SemanticModel model,
        Dictionary<ILocalSymbol, ResultState> states, ReturnStatementSyntax target,
        ref Dictionary<ILocalSymbol, ResultState>? terminal, int depth)
    {
        if (statement is ReturnStatementSyntax)
        {
            if (statement == target) { terminal = states; }
            return null;
        }
        if (statement is BlockSyntax block)
        {
            return AnalyzeStatements(block, model, states, target, ref terminal, depth);
        }
        if (statement is IfStatementSyntax branch && depth < 8)
        {
            InvalidateResultReferences(branch.Condition, model, states);
            var whenTrue = AnalyzeStatement(branch.Statement, model, CloneStates(states), target, ref terminal, depth + 1);
            var whenFalse = branch.Else == null ? states :
                AnalyzeStatement(branch.Else.Statement, model, CloneStates(states), target, ref terminal, depth + 1);
            return MergeStates(whenTrue, whenFalse);
        }
        if (statement is LocalDeclarationStatementSyntax declaration)
        {
            foreach (var variable in declaration.Declaration.Variables)
            {
                if (model.GetDeclaredSymbol(variable) is ILocalSymbol declared)
                {
                    BindResult(declared, variable.Initializer?.Value, model, states);
                }
            }
            return states;
        }
        if (statement is ExpressionStatementSyntax expression)
        {
            if (expression.Expression is AssignmentExpressionSyntax assignment &&
                assignment.IsKind(SyntaxKind.SimpleAssignmentExpression))
            {
                if (ReferencedLocal(assignment.Left, model) is ILocalSymbol assigned)
                {
                    BindResult(assigned, assignment.Right, model, states);
                }
                else if (assignment.Left is MemberAccessExpressionSyntax member &&
                    ReferencedLocal(member.Expression, model) is ILocalSymbol receiver &&
                    states.TryGetValue(receiver, out var state))
                {
                    InvalidateResultReferences(assignment.Right, model, states);
                    SetState(member.Name.Identifier.ValueText, assignment.Right, model, ref state.Success, ref state.Error);
                }
                else { InvalidateResultReferences(assignment, model, states); }
            }
            else { InvalidateResultReferences(statement, model, states); }
            return states;
        }
        if (statement is not EmptyStatementSyntax) { states.Clear(); }
        return states;
    }

    private static Dictionary<ILocalSymbol, ResultState> CloneStates(Dictionary<ILocalSymbol, ResultState> states)
    {
        var copies = new Dictionary<ResultState, ResultState>();
        var clone = new Dictionary<ILocalSymbol, ResultState>(SymbolEqualityComparer.Default);
        foreach (var pair in states)
        {
            if (!copies.TryGetValue(pair.Value, out var copy))
            {
                copy = new ResultState { Success = pair.Value.Success, Error = pair.Value.Error, Escaped = pair.Value.Escaped };
                copies[pair.Value] = copy;
            }
            clone[pair.Key] = copy;
        }
        return clone;
    }

    private static Dictionary<ILocalSymbol, ResultState>? MergeStates(
        Dictionary<ILocalSymbol, ResultState>? left, Dictionary<ILocalSymbol, ResultState>? right)
    {
        if (left == null) { return right; }
        if (right == null) { return left; }
        var merged = new Dictionary<ILocalSymbol, ResultState>(SymbolEqualityComparer.Default);
        var objects = new Dictionary<(ResultState Left, ResultState Right), ResultState>();
        foreach (var pair in left)
        {
            if (!right.TryGetValue(pair.Key, out var other)) { continue; }
            var identity = (pair.Value, other);
            if (!objects.TryGetValue(identity, out var state))
            {
                state = new ResultState
                {
                    Success = pair.Value.Success == other.Success ? other.Success : null,
                    Error = pair.Value.Error == other.Error ? other.Error : null,
                    Escaped = pair.Value.Escaped || other.Escaped
                };
                objects[identity] = state;
            }
            merged[pair.Key] = state;
        }
        // A conditional rebind can make two locals alias on only one path.
        // Their later writes cannot safely update independent merged objects.
        var locals = merged.Keys.ToArray();
        for (var index = 0; index < locals.Length; index++)
        {
            for (var other = index + 1; other < locals.Length; other++)
            {
                var first = locals[index];
                var second = locals[other];
                if (ReferenceEquals(left[first], left[second]) != ReferenceEquals(right[first], right[second]))
                {
                    merged[first].Escaped = true;
                    merged[second].Escaped = true;
                }
            }
        }
        return merged;
    }

    private static ILocalSymbol? ReferencedLocal(ExpressionSyntax expression, SemanticModel model)
    {
        while (expression is ParenthesizedExpressionSyntax parentheses) { expression = parentheses.Expression; }
        var operation = model.GetOperation(expression);
        while (true)
        {
            if (operation is IConversionOperation conversion &&
                (conversion.Conversion.IsIdentity || conversion.Conversion.IsReference))
            {
                operation = conversion.Operand;
            }
            else if (operation is IParenthesizedOperation parenthesized)
            {
                operation = parenthesized.Operand;
            }
            else
            {
                return (operation as ILocalReferenceOperation)?.Local;
            }
        }
    }

    private static void BindResult(ILocalSymbol local, ExpressionSyntax? value, SemanticModel model,
        Dictionary<ILocalSymbol, ResultState> states)
    {
        if (value != null && ReferencedLocal(value, model) is ILocalSymbol source &&
            states.TryGetValue(source, out var shared))
        {
            states[local] = shared;
            return;
        }
        if (value != null) { InvalidateResultReferences(value, model, states); }
        states.Remove(local);
        while (value is ParenthesizedExpressionSyntax or CastExpressionSyntax)
        {
            value = value is ParenthesizedExpressionSyntax parentheses ? parentheses.Expression : ((CastExpressionSyntax)value).Expression;
        }
        if (value == null || !IsResult(model.GetTypeInfo(value).Type)) { return; }

        var state = new ResultState();
        if (value is BaseObjectCreationExpressionSyntax creation && creation.Initializer != null)
        {
            foreach (var assignment in creation.Initializer.Expressions.OfType<AssignmentExpressionSyntax>())
            {
                SetState(assignment.Left.ToString(), assignment.Right, model, ref state.Success, ref state.Error);
            }
        }
        states[local] = state;
    }

    private static void InvalidateResultReferences(SyntaxNode node, SemanticModel model,
        Dictionary<ILocalSymbol, ResultState> states)
    {
        foreach (var identifier in node.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>())
        {
            if (model.GetSymbolInfo(identifier).Symbol is ILocalSymbol referenced &&
                states.TryGetValue(referenced, out var state))
            {
                state.Escaped = true;
            }
        }
    }

    private sealed class ResultState
    {
        public bool? Success;
        public bool? Error;
        public bool Escaped;
    }
}
