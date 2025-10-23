using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Integration tests for ExcelVersionTool
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Version")]
public class ExcelVersionToolTests
{
    [Fact]
    public async Task ExcelVersion_Check_ShouldReturnSuccessJson()
    {
        // Act
        var result = await ExcelVersionTool.ExcelVersion("check");

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        
        // Should have success property
        Assert.True(json.RootElement.TryGetProperty("success", out var successProp));
        
        // If successful, should have version information
        if (successProp.GetBoolean())
        {
            Assert.True(json.RootElement.TryGetProperty("currentVersion", out var currentVersion));
            Assert.NotNull(currentVersion.GetString());
            
            Assert.True(json.RootElement.TryGetProperty("isOutdated", out var isOutdated));
            
            Assert.True(json.RootElement.TryGetProperty("message", out var message));
            Assert.NotNull(message.GetString());
            
            Assert.True(json.RootElement.TryGetProperty("suggestedNextActions", out var actions));
            Assert.True(actions.ValueKind == JsonValueKind.Array);
        }
    }

    [Fact]
    public async Task ExcelVersion_UnknownAction_ShouldThrowException()
    {
        // Act & Assert - Should throw McpException for unknown action
        var exception = await Assert.ThrowsAsync<ModelContextProtocol.McpException>(async () =>
            await ExcelVersionTool.ExcelVersion("unknown"));

        Assert.Contains("Unknown action 'unknown'", exception.Message);
    }

    [Fact]
    public async Task ExcelVersion_Check_ShouldIncludePackageId()
    {
        // Act
        var result = await ExcelVersionTool.ExcelVersion("check");

        // Assert
        var json = JsonDocument.Parse(result);
        
        if (json.RootElement.GetProperty("success").GetBoolean())
        {
            Assert.True(json.RootElement.TryGetProperty("packageId", out var packageId));
            Assert.Equal("Sbroenne.ExcelMcp.McpServer", packageId.GetString());
        }
    }

    [Fact]
    public async Task ExcelVersion_Check_ShouldProvideUpdateCommandWhenOutdated()
    {
        // Act
        var result = await ExcelVersionTool.ExcelVersion("check");

        // Assert
        var json = JsonDocument.Parse(result);
        
        if (json.RootElement.GetProperty("success").GetBoolean() &&
            json.RootElement.GetProperty("isOutdated").GetBoolean())
        {
            Assert.True(json.RootElement.TryGetProperty("updateCommand", out var updateCommand));
            var cmd = updateCommand.GetString();
            Assert.NotNull(cmd);
            Assert.Contains("dotnet tool update", cmd);
            Assert.Contains("Sbroenne.ExcelMcp.McpServer", cmd);
        }
    }

    [Fact]
    public async Task ExcelVersion_Check_ShouldProvideWorkflowHint()
    {
        // Act
        var result = await ExcelVersionTool.ExcelVersion("check");

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("workflowHint", out var workflowHint));
        Assert.NotNull(workflowHint.GetString());
    }
}
