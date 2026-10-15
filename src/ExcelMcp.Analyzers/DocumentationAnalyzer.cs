using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Text;

namespace Sbroenne.ExcelMcp.Analyzers;

/// <summary>Validates advertised counts against the manifest and actual compiled MCP tools.</summary>
[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class DocumentationAnalyzer : DiagnosticAnalyzer
{
    private static readonly DiagnosticDescriptor Count = new(
        "EXL030", "Documentation must match compiled contracts",
        "{0}",
        "Contracts", DiagnosticSeverity.Error, isEnabledByDefault: true,
        customTags: [WellKnownDiagnosticTags.CompilationEnd]);

    private static readonly (string Path, string Pattern)[] Headlines =
    [
        ("README.md", @"(?<m>\d+) MCP tools across (?<t>\d+) feature areas, with (?<o>\d+) operations"),
        ("README.md", @"all (?<o>\d+) operations"),
        ("FEATURES.md", @"(?<t>\d+) feature areas with (?<o>\d+) operations, exposed through (?<m>\d+) MCP tools"),
        ("src/ExcelMcp.McpServer/README.md", @"(?<m>\d+) MCP tools across (?<t>\d+) feature areas, with (?<o>\d+) operations"),
        ("src/ExcelMcp.McpServer/README.md", @"all (?<o>\d+) operations"),
        ("src/ExcelMcp.CLI/README.md", @"provides (?<t>\d+) feature command categories with (?<o>\d+) operations matching"),
        ("src/ExcelMcp.CLI/README.md", @"without loading (?<t>\d+) tool schemas"),
        ("src/ExcelMcp.CLI/README.md", @"\*\*(?<o>\d+) operations\*\* across"),
        ("vscode-extension/README.md", @"(?<m>\d+) MCP tools across (?<t>\d+) feature areas, with (?<o>\d+) operations"),
        ("vscode-extension/README.md", @"all (?<m>\d+) MCP tools and (?<o>\d+) operations"),
        ("mcpb/README.md", @"(?<m>\d+) MCP tools across (?<t>\d+) feature areas, with (?<o>\d+) operations"),
        ("mcpb/manifest.json", @"(?<m>\d+) MCP tools across (?<t>\d+) feature areas, with (?<o>\d+) operations"),
        ("mcpb/BUILD.md", @"generates its (?<m>\d+) tool schemas"),
        ("src/ExcelMcp.CLI/ExcelMcp.CLI.csproj", @"(?<o>\d+) operations across"),
        ("gh-pages/docs/index.md", @"(?<m>\d+) MCP tools and (?<o>\d+) operations"),
        (".github/plugins/excel-mcp/README.md", @"(?<m>\d+) MCP tools across (?<t>\d+) feature areas, with (?<o>\d+) operations"),
        (".github/plugins/excel-cli/README.md", @"command categories with (?<o>\d+) operations"),
        (".github/plugins/excel-cli/README.md", @"\| (?<m>\d+) tool schemas loaded into context \|"),
        ("gh-pages/docs/faq.md", @"same (?<o>\d+) operations"),
        ("docs/INSTALLATION-CLI.md", @"all (?<t>\d+) feature command categories"),
        ("docs/COPILOT-PLUGIN-DISTRIBUTION.md", @"with (?<m>\d+) tools across (?<t>\d+) feature areas \((?<o>\d+) operations\)")
    ];

    private static readonly (string Path, string Pattern)[] SiteLinks =
    [
        ("Directory.Build.props", @"<PackageProjectUrl>https://excelmcpserver\.dev/</PackageProjectUrl>"),
        ("src/ExcelMcp.McpServer/.mcp/server.json", @"""websiteUrl""\s*:\s*""https://excelmcpserver\.dev/"""),
        ("mcpb/manifest.json", @"""homepage""\s*:\s*""https://excelmcpserver\.dev/"""),
        ("vscode-extension/package.json", @"""homepage""\s*:\s*""https://excelmcpserver\.dev/""")
    ];

    /// <inheritdoc />
    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [Count];

    /// <inheritdoc />
    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze | GeneratedCodeAnalysisFlags.ReportDiagnostics);
        context.EnableConcurrentExecution();
        context.RegisterCompilationAction(Analyze);
    }

    private static void Analyze(CompilationAnalysisContext context)
    {
        if (context.Compilation.AssemblyName != "Sbroenne.ExcelMcp.McpServer") { return; }
        var documents = context.Options.AdditionalFiles;
        var countInputs = documents.Where(file => Path.GetFileName(file.Path) == "doc-counts.json").ToArray();
        if (countInputs.Length != 1)
        {
            Report("Missing doc-counts.json compiler input; restore the documentation AdditionalFiles configuration.");
            return;
        }
        var root = countInputs[0].Path.Replace('\\', '/');
        root = root.Substring(0, root.Length - "doc-counts.json".Length);
        var ownTypes = Types(context.Compilation.Assembly.GlobalNamespace).ToArray();
        var core = context.Compilation.SourceModule.ReferencedAssemblySymbols.FirstOrDefault(assembly =>
            assembly.Name == "Sbroenne.ExcelMcp.Core") ?? context.Compilation.Assembly;
        var coreTypes = Types(core.GlobalNamespace).ToArray();
        var json = core.GetTypeByMetadataName("Sbroenne.ExcelMcp.Generated._SkillManifest")?
            .GetMembers("Json").OfType<IFieldSymbol>().FirstOrDefault(field => field.ConstantValue is string)?.ConstantValue as string;
        if (json == null)
        {
            Report("The compiled Core skill manifest is missing; counts cannot be validated.");
            return;
        }
        Manifest manifest;
        try { manifest = Read<Manifest>(json); }
        catch (SerializationException exception)
        {
            Report("The compiled Core skill manifest is invalid: " + exception.Message);
            return;
        }
        if (manifest.Commands == null || manifest.Commands.Any(command => command == null || command.Actions == null) ||
            manifest.TotalCommands <= 0 || manifest.Commands.Length != manifest.TotalCommands ||
            manifest.Commands.Sum(command => (long)command.Actions.Length) != manifest.TotalOperations)
        {
            Report("The compiled Core skill manifest has invalid or inconsistent command and operation totals.");
            return;
        }
        var diagnostics = manifest.Commands.Where(command => command.Name == "diag").ToArray();
        var fileActions = coreTypes.Where(type => type.Name == "FileAction" && type.TypeKind == TypeKind.Enum).ToArray();
        if (diagnostics.Length != 1 || fileActions.Length != 1)
        {
            Report("Expected exactly one CLI-only diag category and one FileAction enum.");
            return;
        }
        var diag = diagnostics[0];
        var fileAction = fileActions[0];
        var fileOperations = fileAction?.GetMembers().OfType<IFieldSymbol>().Count(field =>
            field.GetAttributes().Any(attribute => attribute.AttributeClass?.Name == "JsonStringEnumMemberNameAttribute")) ?? 0;
        if (diag?.Actions == null || fileOperations == 0)
        {
            Report("Expected the CLI-only diag category and the FileAction enum; update count derivation if these contracts intentionally changed.");
            return;
        }
        var areas = manifest.TotalCommands;
        var operations = manifest.TotalOperations - diag.Actions.Length + fileOperations;
        var toolNames = ownTypes.SelectMany(type => type.GetMembers().OfType<IMethodSymbol>())
            .SelectMany(method => method.GetAttributes())
            .Where(attribute => attribute.AttributeClass?.ToDisplayString() == "ModelContextProtocol.Server.McpServerToolAttribute")
            .Select(attribute => attribute.NamedArguments.FirstOrDefault(argument => argument.Key == "Name").Value.Value as string)
            .Where(name => name != null).Select(name => name!).ToArray();
        if (!toolNames.Contains("file", StringComparer.Ordinal) || toolNames.Contains("diag", StringComparer.Ordinal) ||
            toolNames.Where(name => !name.EndsWith("_read", StringComparison.Ordinal)).Distinct(StringComparer.Ordinal).Count() != areas ||
            toolNames.Distinct(StringComparer.Ordinal).Count() != toolNames.Length)
        {
            Report("Compiled MCP endpoints disagree with Core feature areas, contain duplicate names, or expose the CLI-only diag category.");
            return;
        }
        foreach (var (path, pattern) in Headlines)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (Document(path) is not { } document) { continue; }
            if (Matches(document, pattern) is not { } matches) { continue; }
            if (matches.Count == 0) { Report(path + ": expected advertised-count headline is missing.", document); }
            foreach (Match match in matches)
            {
                CheckGroup("t", areas);
                CheckGroup("o", operations);
                CheckGroup("m", toolNames.Length);
                void CheckGroup(string name, int expected)
                {
                    var group = match.Groups[name];
                    if (group.Success && group.Value != expected.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    {
                        Report($"{path}: count is {group.Value}; compiled contracts require {expected}. Update this reviewed headline.",
                            document, new TextSpan(group.Index, group.Length));
                    }
                }
            }
        }
        if (Document("doc-counts.json") is { } counts)
        {
            try
            {
                var advertised = Read<Advertised>(counts.Text.ToString());
                if (advertised.Tools != areas || advertised.Operations != operations)
                {
                    Report($"doc-counts.json must report tools={areas}, operations={operations} from the compiled contracts.", counts);
                }
            }
            catch (SerializationException)
            {
                Report("doc-counts.json must be a JSON object with integer tools and operations properties.", counts);
            }
        }
        var features = documents.Where(file => file.Path.Replace('\\', '/').StartsWith(root + "docs/features/", StringComparison.OrdinalIgnoreCase) &&
            file.Path.EndsWith(".md", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (features.Length == 0) { Report("Canonical feature documents are missing from compiler inputs."); }
        long featureTotal = 0;
        foreach (var file in features)
        {
            if (file.GetText(context.CancellationToken) is not { } text)
            {
                Report("Cannot read canonical feature document: " + file.Path);
                continue;
            }
            var document = new DocumentInput(file.Path, text);
            if (Matches(document, @"(?m)^##\s+.*\((\d+) operations\)") is not { } headings) { continue; }
            foreach (Match match in headings)
            {
                if (!int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture, out var count))
                {
                    Report("Feature-section operation count is not a valid integer.", document);
                    continue;
                }
                featureTotal += count;
            }
        }
        if (featureTotal != operations) { Report($"Feature-section headings total {featureTotal}; compiled contracts require {operations} operations."); }
        foreach (var (path, pattern) in SiteLinks)
        {
            if (Document(path) is { } document && Matches(document, pattern) is { Count: 0 })
            {
                Report(path + ": distribution metadata must link to https://excelmcpserver.dev/.", document);
            }
        }
        return;

        DocumentInput? Document(string relative)
        {
            var candidates = documents.Where(file => string.Equals(file.Path.Replace('\\', '/'),
                root + relative, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (candidates.Length != 1 || candidates[0].GetText(context.CancellationToken) is not { } text)
            {
                Report("Missing or ambiguous compiler document input: " + relative);
                return null;
            }
            return new DocumentInput(candidates[0].Path, text);
        }
        MatchCollection? Matches(DocumentInput document, string pattern)
        {
            try
            {
                var matches = Regex.Matches(document.Text.ToString(), pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
                _ = matches.Count;
                return matches;
            }
            catch (RegexMatchTimeoutException)
            {
                Report("Documentation pattern matching timed out: " + document.Path, document);
                return null;
            }
        }
        void Report(string message, DocumentInput? document = null, TextSpan? span = null)
        {
            var location = document == null ? Location.None : Location.Create(document.Path,
                span ?? default, document.Text.Lines.GetLinePositionSpan(span ?? default));
            context.ReportDiagnostic(Diagnostic.Create(Count, location, message));
        }
    }

    private static IEnumerable<INamedTypeSymbol> Types(INamespaceOrTypeSymbol scope)
    {
        foreach (var member in scope.GetMembers())
        {
            if (member is INamedTypeSymbol type) { yield return type; }
            if (member is INamespaceOrTypeSymbol nested)
            {
                foreach (var descendant in Types(nested)) { yield return descendant; }
            }
        }
    }

    private static T Read<T>(string json)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        return (T)(new DataContractJsonSerializer(typeof(T)).ReadObject(stream)
            ?? throw new SerializationException("Empty JSON document."));
    }

    private sealed class DocumentInput(string path, SourceText text)
    {
        public string Path { get; } = path;
        public SourceText Text { get; } = text;
    }

    [DataContract]
    private sealed class Manifest
    {
        [DataMember(Name = "totalCommands", IsRequired = true)] public int TotalCommands { get; set; }
        [DataMember(Name = "totalOperations", IsRequired = true)] public int TotalOperations { get; set; }
        [DataMember(Name = "commands", IsRequired = true)] public Command[] Commands { get; set; } = [];
    }
    [DataContract]
    private sealed class Command
    {
        [DataMember(Name = "name", IsRequired = true)] public string Name { get; set; } = "";
        [DataMember(Name = "actions", IsRequired = true)] public string[] Actions { get; set; } = [];
    }
    [DataContract]
    private sealed class Advertised
    {
        [DataMember(Name = "tools", IsRequired = true)] public int Tools { get; set; }
        [DataMember(Name = "operations", IsRequired = true)] public int Operations { get; set; }
    }
}
