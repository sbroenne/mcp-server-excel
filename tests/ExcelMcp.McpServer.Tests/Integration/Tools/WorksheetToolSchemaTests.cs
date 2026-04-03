using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Verifies the MCP worksheet tool exposes the intended rename/create parameter contract
/// through the discoverable tool schema.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
public sealed class WorksheetToolSchemaTests : McpIntegrationTestBase
{
    public WorksheetToolSchemaTests(ITestOutputHelper output)
        : base(output, "WorksheetSchemaClient")
    {
    }

    [Fact]
    public async Task ListTools_WorksheetSchema_ExposesActionSpecificWorksheetDescriptions()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var worksheetTool = tools.SingleOrDefault(tool => tool.Name == "worksheet");

        Assert.NotNull(worksheetTool);

        var schema = worksheetTool!.JsonSchema;
        Output.WriteLine(schema.GetRawText());

        var properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("sheet_name", out var sheetNameProperty), $"worksheet schema is missing sheet_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("old_name", out var oldNameProperty), $"worksheet schema is missing old_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("source_name", out var sourceNameProperty), $"worksheet schema is missing source_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("target_name", out var targetNameProperty), $"worksheet schema is missing target_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("new_name", out var newNameProperty), $"worksheet schema is missing new_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("source_file", out var sourceFileProperty), $"worksheet schema is missing source_file: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("source_sheet", out var sourceSheetProperty), $"worksheet schema is missing source_sheet: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("target_file", out var targetFileProperty), $"worksheet schema is missing target_file: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("target_sheet_name", out var targetSheetNameProperty), $"worksheet schema is missing target_sheet_name: {schema.GetRawText()}");

        Assert.Contains("create", GetDescription(sheetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(sheetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rename", GetDescription(oldNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rename", GetDescription(newNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy", GetDescription(sourceNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(sourceNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy", GetDescription(targetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(targetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(sourceFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move-to-file", GetDescription(sourceFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(sourceSheetProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move-to-file", GetDescription(sourceSheetProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(sourceSheetProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(targetFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move-to-file", GetDescription(targetFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(targetSheetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(targetSheetNameProperty), StringComparison.OrdinalIgnoreCase);
    }

    private static string GetDescription(JsonElement property)
    {
        Assert.True(property.TryGetProperty("description", out var description), $"Schema property is missing description: {property.GetRawText()}");
        return description.GetString() ?? string.Empty;
    }
}
