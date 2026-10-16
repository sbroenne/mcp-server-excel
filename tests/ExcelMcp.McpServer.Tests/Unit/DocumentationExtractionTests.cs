using System.Reflection;
using System.Xml;
using Microsoft.CodeAnalysis;
using Sbroenne.ExcelMcp.Generators.Common;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "GeneratedContracts")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DocumentationExtractionTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MalformedDocumentation_FailsWithSymbolContextAndOriginalXmlError(bool interfaceDocumentation)
    {
        var symbol = interfaceDocumentation
            ? (ISymbol)DispatchProxy.Create<INamedTypeSymbol, DocumentationSymbolProxy>()
            : DispatchProxy.Create<IMethodSymbol, DocumentationSymbolProxy>();
        var method = typeof(ServiceInfoExtractor).GetMethod(
            interfaceDocumentation ? "ExtractInterfaceSummary" : "ExtractXmlDocumentation",
            BindingFlags.NonPublic | BindingFlags.Static)!;

        var invocation = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, [symbol]));
        var failure = Assert.IsType<InvalidOperationException>(invocation.InnerException);
        Assert.Contains("SampleContract", failure.Message, StringComparison.Ordinal);
        var xmlFailure = Assert.IsType<XmlException>(failure.InnerException);
        Assert.Contains(xmlFailure.Message, failure.Message, StringComparison.Ordinal);
    }

    public class DocumentationSymbolProxy : DispatchProxy
    {
        protected override object? Invoke(System.Reflection.MethodInfo? targetMethod, object?[]? args) =>
            targetMethod?.Name switch
            {
                "GetDocumentationCommentXml" => "<summary>Unterminated documentation",
                "ToDisplayString" => "SampleContract",
                _ => throw new NotSupportedException(targetMethod?.Name)
            };
    }
}
