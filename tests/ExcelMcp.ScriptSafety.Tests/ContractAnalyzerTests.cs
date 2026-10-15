using Sbroenne.ExcelMcp.Analyzers;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Feature", "PreCommit")]
[Trait("RequiresExcel", "false")]
public sealed class ContractAnalyzerTests
{
    private const string Attributes = """
        using System;
        using System.Threading;
        using System.Threading.Tasks;
        namespace Sbroenne.ExcelMcp.Core.Attributes {
            class ServiceCategoryAttribute : Attribute {}
        }
        namespace ModelContextProtocol.Server {
            class McpServerToolAttribute : Attribute {}
        }
        namespace ModelContextProtocol.Protocol { class CallToolResult {} }
        namespace Other {
            class ServiceCategoryAttribute : Attribute {}
            class CallToolResult {}
        }
        """;

    [Theory]
    [InlineData("Task")]
    [InlineData("Task<int>")]
    [InlineData("ValueTask")]
    [InlineData("ValueTask<int>")]
    public async Task CoreContract_RejectsAsyncResults(string result)
    {
        var diagnostics = await AnalyzeAsync($$"""
            [Sbroenne.ExcelMcp.Core.Attributes.ServiceCategory]
            interface Commands { {{result}} Read(); }
            """);
        Assert.Equal("EXL013", Assert.Single(diagnostics).Id);
    }

    [Theory]
    [InlineData("workbookSessionId", false)]
    [InlineData("xmlText", false)]
    [InlineData("value2", false)]
    [InlineData("WorkbookSessionId", true)]
    [InlineData("workbook_session_id", true)]
    [InlineData("_value", true)]
    public async Task CoreContract_UsesActualCamelCase(string parameter, bool invalid)
    {
        var diagnostics = await AnalyzeAsync($$"""
            [Sbroenne.ExcelMcp.Core.Attributes.ServiceCategory]
            interface Commands { int Read(string {{parameter}}); }
            """);
        if (invalid) { Assert.Equal("EXL016", Assert.Single(diagnostics).Id); }
        else { Assert.Empty(diagnostics); }
    }

    [Fact]
    public async Task UnrelatedAttribute_DoesNotImposeCoreContractRules()
    {
        Assert.Empty(await AnalyzeAsync("""
            [Other.ServiceCategory]
            interface Commands { Task Read(string NonCoreName); }
            """));
    }

    [Theory]
    [InlineData("Task<ModelContextProtocol.Protocol.CallToolResult>", "CancellationToken cancellationToken", false)]
    [InlineData("Task<ModelContextProtocol.Protocol.CallToolResult>", "CancellationToken cancellationToken = default", false)]
    [InlineData("Task<ModelContextProtocol.Protocol.CallToolResult>", "", true)]
    [InlineData("Task<Other.CallToolResult>", "CancellationToken cancellationToken", true)]
    [InlineData("ValueTask<ModelContextProtocol.Protocol.CallToolResult>", "CancellationToken cancellationToken", true)]
    [InlineData("string", "CancellationToken cancellationToken", true)]
    public async Task McpContract_RequiresSdkResultAndCancellation(string result, string parameters, bool invalid)
    {
        var diagnostics = await AnalyzeAsync($$"""
            abstract class Tools {
                [ModelContextProtocol.Server.McpServerTool]
                public abstract {{result}} Read({{parameters}});
            }
            """);
        if (invalid) { Assert.Equal("EXL014", Assert.Single(diagnostics).Id); }
        else { Assert.Empty(diagnostics); }
    }

    private static Task<System.Collections.Immutable.ImmutableArray<Microsoft.CodeAnalysis.Diagnostic>> AnalyzeAsync(string source) =>
        SafetyAnalyzerTests.AnalyzeAsync(Attributes + source, new ContractAnalyzer());
}
