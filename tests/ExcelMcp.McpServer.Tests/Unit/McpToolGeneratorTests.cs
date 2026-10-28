using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Sbroenne.ExcelMcp.Generators.Common;
using Sbroenne.ExcelMcp.Generators.Mcp;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "McpToolGeneration")]
[Trait("Feature", "GeneratedContracts")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class McpToolGeneratorTests
{
    [Fact]
    public void WorksheetAndScreenshot_AreGeneratedFromCoreContracts()
    {
        Assert.Equal("worksheet", Sbroenne.ExcelMcp.Generated.ServiceRegistry.Sheet.McpToolName);
        Assert.Equal("sheet", Sbroenne.ExcelMcp.Generated.ServiceRegistry.Sheet.Category);
        Assert.Equal("sheet", Sbroenne.ExcelMcp.Generated.ServiceRegistry.Sheet.CliCommandName);
        Assert.Equal("screenshot", Sbroenne.ExcelMcp.Generated.ServiceRegistry.Screenshot.CliCommandName);
        Assert.NotNull(typeof(Tools.ExcelSheet_worksheetTool).GetMethod("ExcelSheet_worksheet"));
        Assert.NotNull(typeof(Tools.ExcelSheet_worksheet_005FreadTool).GetMethod("ExcelSheet_worksheet_005Fread"));
        Assert.NotNull(typeof(Tools.ExcelScreenshot_screenshotTool).GetMethod("ExcelScreenshot_screenshot"));
    }

    [Fact]
    public void MixedSessionActions_GenerateOptionalWriteSessionAndRequiredReadSession()
    {
        var generated = Generate("""
            using System;
            public sealed class ServiceCategoryAttribute(string name) : Attribute;
            public sealed class McpToolAttribute(string name) : Attribute;
            public sealed class McpReadOnlyActionsAttribute(params string[] actions) : Attribute;
            public interface IExcelBatch;
            [ServiceCategory("Sheet"), McpTool("worksheet"), McpReadOnlyActions("list")]
            public interface ISheetCommands
            {
                string List(IExcelBatch batch);
                string Create(IExcelBatch batch, string sheetName);
                string CopyToFile(string sourceFile);
            }
            """).GeneratedSources.Select(item => item.SourceText.ToString()).ToArray();

        var write = Assert.Single(generated, source => source.Contains("Name = \"worksheet\""));
        var read = Assert.Single(generated, source => source.Contains("Name = \"worksheet_read\""));
        Assert.Contains("string? workbook_session_id = null", write);
        Assert.Contains("[McpMeta(\"requiresSession\", false)]", write);
        Assert.Contains("string workbook_session_id,", read);
        Assert.Contains("[McpMeta(\"requiresSession\", true)]", read);
        Assert.Contains("SheetAction.CopyToFile", write);
        Assert.Contains("SheetAction.Create", write);
    }

    [Fact]
    public void ScreenshotResult_GeneratesImageHelperWithoutOutputSchema()
    {
        var generated = Generate("""
            using System;
            public sealed class ServiceCategoryAttribute(string name) : Attribute;
            public sealed class McpToolAttribute(string name) : Attribute;
            public interface IExcelBatch;
            namespace Sbroenne.ExcelMcp.Core.Models
            {
                public class ResultBase { public bool Success { get; set; } }
            }
            namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot
            {
                public class ScreenshotResult : Sbroenne.ExcelMcp.Core.Models.ResultBase
                {
                    public string ImageBase64 { get; set; }
                    public string MimeType { get; set; }
                    public int Width { get; set; }
                }
                [ServiceCategory("Screenshot"), McpTool("screenshot")]
                public interface IScreenshotCommands
                {
                    ScreenshotResult Capture(IExcelBatch batch);
                }
            }
            """).GeneratedSources.Select(item => item.SourceText.ToString()).ToArray();

        var screenshot = Assert.Single(generated);
        Assert.Contains("ScreenshotImageResult.CreateToolResult", screenshot);
        Assert.DoesNotContain("OutputSchema", screenshot);
        Assert.DoesNotContain("UseStructuredContent", screenshot);
    }

    private static GeneratorRunResult Generate(string source)
    {
        var runtimeDirectory = Path.GetDirectoryName(typeof(object).Assembly.Location)!;
        var references = new[]
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(Path.Combine(runtimeDirectory, "System.Runtime.dll"))
        };
        var contracts = CSharpCompilation.Create(
            "GeneratorTestContracts",
            [CSharpSyntaxTree.ParseText(source)],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using var assembly = new MemoryStream();
        var emitted = contracts.Emit(assembly);
        Assert.True(emitted.Success, string.Join(Environment.NewLine, emitted.Diagnostics));
        var compilation = CSharpCompilation.Create(
            "GeneratorTestServer",
            references: references.Append(MetadataReference.CreateFromImage(assembly.ToArray())),
            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        var result = CSharpGeneratorDriver.Create(new McpToolGenerator().AsSourceGenerator())
            .RunGenerators(compilation).GetRunResult();
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DiagnosticSeverity.Error);
        var generator = Assert.Single(result.Results);
        Assert.Null(generator.Exception);
        return generator;
    }

    [Fact]
    public void SharedToolContract_ProjectsTimeoutsAndRequirednessOnce()
    {
        var source = """
            using System;
            public sealed class ServiceCategoryAttribute(string name) : Attribute;
            public sealed class McpToolAttribute(string name) : Attribute;
            public sealed class RequiredParameterAttribute : Attribute;
            [ServiceCategory("sample"), McpTool("sample")]
            public interface ISampleCommands
            {
                string Refresh([RequiredParameter] int? limit = null, TimeSpan? timeout = null);
            }
            """;
        var runtimeDirectory = Path.GetDirectoryName(typeof(object).Assembly.Location)!;
        var references = new[]
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(Path.Combine(runtimeDirectory, "System.Runtime.dll"))
        };
        var compilation = CSharpCompilation.Create(
            "SharedToolContractTests",
            [CSharpSyntaxTree.ParseText(source)],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        var interfaceSymbol = compilation.GetTypeByMetadataName("ISampleCommands");
        Assert.NotNull(interfaceSymbol);

        var serviceInfo = ServiceInfoExtractor.ExtractServiceInfo(interfaceSymbol);
        Assert.NotNull(serviceInfo);
        var parameters = ServiceInfoExtractor.GetAllExposedParameters(serviceInfo);

        var limit = Assert.Single(parameters, parameter => parameter.Name == "limit");
        Assert.Equal("int?", limit.TypeName);
        Assert.Contains("refresh", limit.RequiredByActions);
        Assert.Contains("refresh", limit.ApplicableByActions);

        var timeout = Assert.Single(parameters, parameter => parameter.Name == "timeoutSeconds");
        Assert.Equal("int?", timeout.TypeName);
        Assert.Empty(timeout.RequiredByActions);
        Assert.Contains("refresh", timeout.ApplicableByActions);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MethodToolOverrides_GenerateDistinctSourcesAndDeclarations(bool readOnly)
    {
        var source = $$"""
            using System;
            public sealed class ServiceCategoryAttribute(string name) : Attribute;
            public sealed class McpToolAttribute(string name) : Attribute
            {
                public bool ReadOnly { get; set; }
            }
            public sealed class NoSessionAttribute : Attribute;
            [ServiceCategory("sample"), McpTool("sample"), NoSession]
            public interface ISampleCommands
            {
                [McpTool("sample", ReadOnly = {{readOnly.ToString().ToLowerInvariant()}})]
                string First();
                [McpTool("sample_other", ReadOnly = {{readOnly.ToString().ToLowerInvariant()}})]
                string Second();
                [McpTool("sample-other", ReadOnly = {{readOnly.ToString().ToLowerInvariant()}})]
                string Third();
            }
            """;
        var generated = Generate(source).GeneratedSources;
        Assert.Equal(3, generated.Length);
        Assert.Equal(3, generated.Select(item => item.HintName).Distinct(StringComparer.Ordinal).Count());
        var types = generated.SelectMany(item => item.SyntaxTree.GetRoot()
            .DescendantNodes().OfType<BaseTypeDeclarationSyntax>()
            .Select(type => type.Identifier.ValueText)).ToArray();
        Assert.Equal(6, types.Length);
        Assert.Equal(types.Length, types.Distinct(StringComparer.Ordinal).Count());
        var methods = generated.SelectMany(item => item.SyntaxTree.GetRoot()
            .DescendantNodes().OfType<MethodDeclarationSyntax>()
            .Select(method => method.Identifier.ValueText)).ToArray();
        Assert.Equal(3, methods.Distinct(StringComparer.Ordinal).Count());
        foreach (var toolName in new[] { "sample", "sample_other", "sample-other" })
        {
            Assert.Single(generated.Where(item => item.SourceText.ToString()
                .Contains($"Name = \"{toolName}\"", StringComparison.Ordinal)));
        }
    }
}
