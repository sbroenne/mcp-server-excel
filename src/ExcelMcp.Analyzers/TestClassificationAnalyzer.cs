using System;
using System.Collections.Concurrent;
using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

namespace Sbroenne.ExcelMcp.Analyzers;

/// <summary>Rejects unclassified or parallel Excel tests before a test host starts.</summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class TestClassificationAnalyzer : DiagnosticAnalyzer
{
    private static readonly DiagnosticDescriptor Classification = new(
        "EXL020", "Test needs one Excel classification",
        "'{0}' must have exactly one distinct RequiresExcel=true/false trait on its method or class",
        "TestSafety", DiagnosticSeverity.Error, isEnabledByDefault: true, customTags: [WellKnownDiagnosticTags.CompilationEnd]);
    private static readonly DiagnosticDescriptor Exclusive = new(
        "EXL021", "Excel tests must not run in parallel",
        "'{0}' needs a Collection whose CollectionDefinition has DisableParallelization=true",
        "TestSafety", DiagnosticSeverity.Error, isEnabledByDefault: true, customTags: [WellKnownDiagnosticTags.CompilationEnd]);
    private static readonly DiagnosticDescriptor MixedFixtures = new(
        "EXL022", "Do not combine class and collection fixtures",
        "'{0}' has both a class fixture and a collection fixture; choose one owner for shared state",
        "TestSafety", DiagnosticSeverity.Error, isEnabledByDefault: true, customTags: [WellKnownDiagnosticTags.CompilationEnd]);
    private static readonly DiagnosticDescriptor SystemTestOwner = new(
        "EXL023", "System test needs an owning process-test group",
        "'{0}' uses AdapterTestKind=System outside the CLI test project; add its owner to the hosted Process selection first",
        "TestSafety", DiagnosticSeverity.Error, isEnabledByDefault: true, customTags: [WellKnownDiagnosticTags.CompilationEnd]);

    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [Classification, Exclusive, MixedFixtures, SystemTestOwner];

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterCompilationStartAction(start =>
        {
            var collections = new ConcurrentDictionary<string, (bool Exclusive, bool Fixture)>(StringComparer.Ordinal);
            var tests = new ConcurrentBag<IMethodSymbol>();
            start.RegisterSymbolAction(symbolContext =>
            {
                var type = (INamedTypeSymbol)symbolContext.Symbol;
                foreach (var definition in type.GetAttributes().Where(a =>
                    a.AttributeClass?.ToDisplayString() == "Xunit.CollectionDefinitionAttribute"))
                {
                    if (definition.ConstructorArguments.FirstOrDefault().Value is not string name) { continue; }
                    collections[name] = (
                        definition.NamedArguments.Any(a => a.Key == "DisableParallelization" && a.Value.Value is true),
                        type.AllInterfaces.Any(i => i.MetadataName == "ICollectionFixture`1" &&
                            i.ContainingNamespace.ToDisplayString() == "Xunit"));
                }
            }, SymbolKind.NamedType);
            start.RegisterSymbolAction(symbolContext =>
            {
                var method = (IMethodSymbol)symbolContext.Symbol;
                if (method.GetAttributes().Any(a => IsFact(a.AttributeClass))) { tests.Add(method); }
            }, SymbolKind.Method);
            start.RegisterCompilationEndAction(end =>
            {
                foreach (var method in tests)
                {
                    var type = method.ContainingType;
                    var attributes = type.GetAttributes().Concat(method.GetAttributes()).ToArray();
                    if (end.Compilation.AssemblyName != "Sbroenne.ExcelMcp.CLI.Tests" &&
                        attributes.Any(a => a.AttributeClass?.ToDisplayString() == "Xunit.TraitAttribute" &&
                            a.ConstructorArguments.Length == 2 &&
                            a.ConstructorArguments[0].Value as string == "AdapterTestKind" &&
                            a.ConstructorArguments[1].Value as string == "System"))
                    {
                        end.ReportDiagnostic(Diagnostic.Create(SystemTestOwner, method.Locations.FirstOrDefault(), method.Name));
                    }
                    var classifications = attributes
                        .Where(a => a.AttributeClass?.ToDisplayString() == "Xunit.TraitAttribute" &&
                            a.ConstructorArguments.Length == 2 &&
                            string.Equals(a.ConstructorArguments[0].Value as string, "RequiresExcel", StringComparison.OrdinalIgnoreCase))
                        .Select(a => (a.ConstructorArguments[1].Value as string)?.ToLowerInvariant())
                        .Distinct(StringComparer.Ordinal).ToArray();
                    if (classifications.Length != 1 || classifications[0] is not ("true" or "false"))
                    {
                        end.ReportDiagnostic(Diagnostic.Create(Classification, method.Locations.FirstOrDefault(), method.Name));
                        continue;
                    }
                    var collection = type.GetAttributes().FirstOrDefault(a =>
                        a.AttributeClass?.ToDisplayString() == "Xunit.CollectionAttribute");
                    var name = collection?.ConstructorArguments.FirstOrDefault().Value as string;
                    var definition = name != null && collections.TryGetValue(name, out var value) ? value : default;
                    if (classifications[0] == "true" && !definition.Exclusive)
                    {
                        end.ReportDiagnostic(Diagnostic.Create(Exclusive, method.Locations.FirstOrDefault(), method.Name));
                    }
                    if (definition.Fixture &&
                        type.AllInterfaces.Any(i => i.MetadataName == "IClassFixture`1" &&
                            i.ContainingNamespace.ToDisplayString() == "Xunit"))
                    {
                        end.ReportDiagnostic(Diagnostic.Create(MixedFixtures, method.Locations.FirstOrDefault(), method.Name));
                    }
                }
            });
        });
    }

    private static bool IsFact(INamedTypeSymbol? type)
    {
        for (; type != null; type = type.BaseType)
        {
            if (type.ToDisplayString() == "Xunit.FactAttribute") { return true; }
        }
        return false;
    }
}
