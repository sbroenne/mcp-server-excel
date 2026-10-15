using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Sbroenne.ExcelMcp.Analyzers;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Feature", "PreCommit")]
[Trait("RequiresExcel", "false")]
public sealed class DocumentationAnalyzerTests
{
    private const string Manifest = """
        {"totalCommands":1,"totalOperations":1,"commands":[{"name":"diag","actions":["self-test"]}]}
        """;

    [Fact]
    public async Task MatchingDocumentation_AcceptsCompiledCounts()
    {
        Assert.Empty(await AnalyzeAsync(Documents()));
    }

    [Theory]
    [InlineData("README.md", "2 MCP tools across 1 feature areas, with 1 operations\nall 1 operations")]
    [InlineData("doc-counts.json", """{"tools":1,"operations":2}""")]
    [InlineData("doc-counts.json", "not JSON")]
    [InlineData("doc-counts.json", """{"tools":null,"operations":1}""")]
    [InlineData("docs\\features\\files.md", "## Files (2 operations)")]
    [InlineData("docs\\features\\files.md", "## Files (99999999999999999999999 operations)")]
    [InlineData("vscode-extension\\package.json", """{"homepage":"https://example.com/"}""")]
    public async Task InvalidDocumentation_ReportsAnErrorNotAnAnalyzerCrash(string path, string replacement)
    {
        var documents = Documents();
        documents[path] = replacement;
        var diagnostics = await AnalyzeAsync(documents);
        Assert.NotEmpty(diagnostics);
        Assert.All(diagnostics, diagnostic => Assert.Equal("EXL030", diagnostic.Id));
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("null")]
    [InlineData("""{"totalCommands":1,"totalOperations":1,"commands":null}""")]
    [InlineData("""{"totalCommands":1,"totalOperations":1,"commands":[{"name":"diag","actions":[]},{"name":"diag","actions":[]}]}""")]
    public async Task InvalidManifest_ReportsAnErrorNotAnAnalyzerCrash(string manifest)
    {
        var diagnostics = await AnalyzeAsync(Documents(), manifest);
        Assert.NotEmpty(diagnostics);
        Assert.All(diagnostics, diagnostic => Assert.Equal("EXL030", diagnostic.Id));
    }

    [Fact]
    public async Task MissingCompilerInput_FailsExplicitly()
    {
        var documents = Documents();
        documents.Remove("FEATURES.md");
        Assert.Contains(await AnalyzeAsync(documents), diagnostic =>
            diagnostic.Id == "EXL030" && diagnostic.GetMessage(System.Globalization.CultureInfo.InvariantCulture)
                .Contains("FEATURES.md", StringComparison.Ordinal));
    }

    private static async Task<ImmutableArray<Diagnostic>> AnalyzeAsync(Dictionary<string, string> documents,
        string manifest = Manifest)
    {
        var source = """
            using System;
            using System.Text.Json.Serialization;
            namespace ModelContextProtocol.Server {
                class McpServerToolAttribute : Attribute { public string Name { get; set; } = ""; }
            }
            class Tools {
                [ModelContextProtocol.Server.McpServerTool(Name = "file")] public void File() {}
            }
            enum FileAction { [JsonStringEnumMemberName("open")] Open }
            namespace Sbroenne.ExcelMcp.Generated { class _SkillManifest { public const string Json =
            """ + System.Text.Json.JsonSerializer.Serialize(manifest) + "; } }";
        return await SafetyAnalyzerTests.AnalyzeAsync(source, new DocumentationAnalyzer(),
            "Sbroenne.ExcelMcp.McpServer",
            documents.Select(pair => (AdditionalText)new Document(pair.Key, pair.Value)).ToImmutableArray());
    }

    private static Dictionary<string, string> Documents()
    {
        const string headline = "1 MCP tools across 1 feature areas, with 1 operations";
        return new(StringComparer.Ordinal)
        {
            ["doc-counts.json"] = """{"tools":1,"operations":1}""",
            ["README.md"] = headline + "\nall 1 operations",
            ["FEATURES.md"] = "1 feature areas with 1 operations, exposed through 1 MCP tools",
            ["src\\ExcelMcp.McpServer\\README.md"] = headline + "\nall 1 operations",
            ["src\\ExcelMcp.CLI\\README.md"] = "provides 1 feature command categories with 1 operations matching\nwithout loading 1 tool schemas\n**1 operations** across",
            ["src\\ExcelMcp.CLI\\ExcelMcp.CLI.csproj"] = "1 operations across",
            ["vscode-extension\\README.md"] = headline + "\nall 1 MCP tools and 1 operations",
            ["vscode-extension\\package.json"] = """{"homepage":"https://excelmcpserver.dev/"}""",
            ["mcpb\\README.md"] = headline,
            ["mcpb\\manifest.json"] = """{"homepage":"https://excelmcpserver.dev/"}""" + headline,
            ["mcpb\\BUILD.md"] = "generates its 1 tool schemas",
            ["gh-pages\\docs\\index.md"] = "1 MCP tools and 1 operations",
            ["gh-pages\\docs\\faq.md"] = "same 1 operations",
            [".github\\plugins\\excel-mcp\\README.md"] = headline,
            [".github\\plugins\\excel-cli\\README.md"] = "command categories with 1 operations\n| 1 tool schemas loaded into context |",
            ["docs\\INSTALLATION-CLI.md"] = "all 1 feature command categories",
            ["docs\\COPILOT-PLUGIN-DISTRIBUTION.md"] = "with 1 tools across 1 feature areas (1 operations)",
            ["Directory.Build.props"] = "<PackageProjectUrl>https://excelmcpserver.dev/</PackageProjectUrl>",
            ["src\\ExcelMcp.McpServer\\.mcp\\server.json"] = """{"websiteUrl":"https://excelmcpserver.dev/"}""",
            ["docs\\features\\files.md"] = "## Files (1 operations)"
        };
    }

    private sealed class Document(string path, string text) : AdditionalText
    {
        public override string Path => "C:\\Fixture\\" + path;
        public override SourceText GetText(CancellationToken cancellationToken = default) => SourceText.From(text);
    }
}
