using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit.Commands;

/// <summary>
/// Unit tests for PowerQueryWorkflowGuidance batch mode suggestions.
/// Tests verify that batch mode hints are provided appropriately.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "WorkflowGuidance")]
public class PowerQueryWorkflowGuidanceBatchModeTests
{
    [Fact]
    public void GetNextStepsAfterImport_WithoutBatchMode_SuggestsBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterImport(
            isConnectionOnly: false,
            hasErrors: false,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // First suggestion should be about batch mode
        Assert.Contains(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(suggestions, s => s.Contains("multiple imports", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterImport_WithBatchMode_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterImport(
            isConnectionOnly: false,
            hasErrors: false,
            usedBatchMode: true);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should NOT suggest batch mode since already using it
        Assert.DoesNotContain(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterImport_WithErrors_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterImport(
            isConnectionOnly: false,
            hasErrors: true,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Error scenario - should focus on fixing errors, not batch mode
        Assert.DoesNotContain(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(suggestions, s => s.Contains("error", StringComparison.OrdinalIgnoreCase) || 
                                          s.Contains("fix", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterUpdate_WithoutBatchMode_SuggestsBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
            configPreserved: true,
            hasErrors: false,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should suggest batch mode for multiple updates
        Assert.Contains(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(suggestions, s => s.Contains("multiple updates", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterUpdate_WithBatchMode_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
            configPreserved: true,
            hasErrors: false,
            usedBatchMode: true);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should NOT suggest batch mode since already using it
        Assert.DoesNotContain(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterLoadConfig_WithoutBatchMode_SuggestsBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterLoadConfig(
            loadMode: "LoadToTable",
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should suggest batch mode for configuring multiple queries
        Assert.Contains(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(suggestions, s => s.Contains("multiple queries", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterLoadConfig_WithBatchMode_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterLoadConfig(
            loadMode: "LoadToDataModel",
            usedBatchMode: true);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should NOT suggest batch mode since already using it
        Assert.DoesNotContain(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(false, 3)] // Without batch mode, should have at least 3 suggestions
    [InlineData(true, 2)]  // With batch mode, may have fewer suggestions (no batch hint)
    public void GetNextStepsAfterImport_HasAppropriateNumberOfSuggestions(bool usedBatchMode, int minimumCount)
    {
        // Arrange & Act
        var suggestions = PowerQueryWorkflowGuidance.GetNextStepsAfterImport(
            isConnectionOnly: false,
            hasErrors: false,
            usedBatchMode: usedBatchMode);

        // Assert
        Assert.NotNull(suggestions);
        Assert.True(suggestions.Count >= minimumCount, 
            $"Expected at least {minimumCount} suggestions, got {suggestions.Count}");
    }
}
