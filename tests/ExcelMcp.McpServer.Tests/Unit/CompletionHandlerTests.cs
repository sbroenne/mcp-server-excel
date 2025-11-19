using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.McpServer.Completions;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

/// <summary>
/// Tests for ExcelCompletionHandler autocomplete functionality
/// </summary>
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
public class CompletionHandlerTests
{
    [Fact]
    public void HandleCompletion_TableStyleParameter_ReturnsTableStyleCompletions()
    {
        // Arrange
        var request = new JsonObject
        {
            ["params"] = new JsonObject
            {
                ["ref"] = new JsonObject
                {
                    ["type"] = "ref/prompt"
                },
                ["argument"] = new JsonObject
                {
                    ["name"] = "tableStyle",
                    ["value"] = ""
                }
            }
        };

        // Act
        var result = ExcelCompletionHandler.HandleCompletion(request);

        // Assert
        Assert.NotNull(result);
        var values = result["values"] as JsonArray;
        Assert.NotNull(values);
        Assert.True(values.Count > 0, "Should return table style completions");

        // Verify some expected styles are present
        var valueStrings = values.Select(v => v?["value"]?.ToString()).ToList();
        Assert.Contains("TableStyleMedium2", valueStrings);
        Assert.Contains("TableStyleLight9", valueStrings);
        Assert.Contains("TableStyleDark1", valueStrings);
    }

    [Fact]
    public void HandleCompletion_TableStyleParameter_WithPrefix_ReturnsFilteredCompletions()
    {
        // Arrange
        var request = new JsonObject
        {
            ["params"] = new JsonObject
            {
                ["ref"] = new JsonObject
                {
                    ["type"] = "ref/prompt"
                },
                ["argument"] = new JsonObject
                {
                    ["name"] = "tableStyle",
                    ["value"] = "TableStyleMedium"
                }
            }
        };

        // Act
        var result = ExcelCompletionHandler.HandleCompletion(request);

        // Assert
        Assert.NotNull(result);
        var values = result["values"] as JsonArray;
        Assert.NotNull(values);
        Assert.True(values.Count > 0, "Should return filtered table style completions");

        // Verify all results start with the prefix
        var valueStrings = values.Select(v => v?["value"]?.ToString()).ToList();
        Assert.All(valueStrings, v =>
            Assert.StartsWith("TableStyleMedium", v, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void HandleCompletion_LoadDestinationParameter_ReturnsLoadDestinationCompletions()
    {
        // Arrange
        var request = new JsonObject
        {
            ["params"] = new JsonObject
            {
                ["ref"] = new JsonObject
                {
                    ["type"] = "ref/prompt"
                },
                ["argument"] = new JsonObject
                {
                    ["name"] = "loadDestination",
                    ["value"] = ""
                }
            }
        };

        // Act
        var result = ExcelCompletionHandler.HandleCompletion(request);

        // Assert
        Assert.NotNull(result);
        var values = result["values"] as JsonArray;
        Assert.NotNull(values);
        Assert.True(values.Count > 0, "Should return load destination completions");
    }

    [Fact]
    public void HandleCompletion_UnknownParameter_ReturnsEmptyCompletions()
    {
        // Arrange
        var request = new JsonObject
        {
            ["params"] = new JsonObject
            {
                ["ref"] = new JsonObject
                {
                    ["type"] = "ref/prompt"
                },
                ["argument"] = new JsonObject
                {
                    ["name"] = "unknownParameter",
                    ["value"] = ""
                }
            }
        };

        // Act
        var result = ExcelCompletionHandler.HandleCompletion(request);

        // Assert
        Assert.NotNull(result);
        var values = result["values"] as JsonArray;
        Assert.NotNull(values);
        Assert.Empty(values);
    }
}
