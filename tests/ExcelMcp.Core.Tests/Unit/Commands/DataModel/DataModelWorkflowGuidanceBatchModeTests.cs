using Sbroenne.ExcelMcp.Core.Commands.DataModel;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit.Commands.DataModel;

/// <summary>
/// Unit tests for DataModelWorkflowGuidance batch mode suggestions.
/// Tests verify that batch mode hints are provided appropriately for Data Model operations.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "WorkflowGuidance")]
public class DataModelWorkflowGuidanceBatchModeTests
{
    [Fact]
    public void GetNextStepsAfterCreateMeasure_WithoutBatchMode_SuggestsBatchMode()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterCreateMeasure(
            success: true,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // First suggestion should be about batch mode
        Assert.Contains(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase) ||
                                          s.Contains("batch mode", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(suggestions, s => s.Contains("multiple measures", StringComparison.OrdinalIgnoreCase) ||
                                          s.Contains("Creating multiple", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterCreateMeasure_WithBatchMode_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterCreateMeasure(
            success: true,
            usedBatchMode: true);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should NOT suggest batch mode since already using it
        Assert.DoesNotContain(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterCreateMeasure_WithError_ProvidesRecoverySteps()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterCreateMeasure(
            success: false,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Error scenario - should focus on fixing errors
        Assert.Contains(suggestions, s => s.Contains("failed", StringComparison.OrdinalIgnoreCase) ||
                                          s.Contains("error", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterCreateRelationship_WithoutBatchMode_SuggestsBatchMode()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterCreateRelationship(
            success: true,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should suggest batch mode for multiple relationships
        Assert.Contains(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase) ||
                                          s.Contains("batch mode", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(suggestions, s => s.Contains("multiple relationships", StringComparison.OrdinalIgnoreCase) ||
                                          s.Contains("Creating multiple", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterCreateRelationship_WithBatchMode_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterCreateRelationship(
            success: true,
            usedBatchMode: true);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should NOT suggest batch mode since already using it
        Assert.DoesNotContain(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterCreateColumn_WithoutBatchMode_SuggestsBatchMode()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterCreateColumn(
            success: true,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        
        // Should suggest batch mode for multiple columns
        Assert.Contains(suggestions, s => s.Contains("begin_excel_batch", StringComparison.OrdinalIgnoreCase) ||
                                          s.Contains("batch mode", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetWorkflowHint_ForMultiOperationScenario_MentionsBatchMode()
    {
        // Arrange & Act
        var hint = DataModelWorkflowGuidance.GetWorkflowHint(
            operation: "create-measure",
            success: true,
            usedBatchMode: false);

        // Assert
        Assert.NotNull(hint);
        Assert.Contains("batch mode", hint, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetWorkflowHint_WhenUsingBatchMode_DoesNotSuggestBatchMode()
    {
        // Arrange & Act
        var hint = DataModelWorkflowGuidance.GetWorkflowHint(
            operation: "create-measure",
            success: true,
            usedBatchMode: true);

        // Assert
        Assert.NotNull(hint);
        // Should not suggest batch mode when already using it
        Assert.DoesNotContain("consider", hint, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetNextStepsAfterList_WithZeroItems_SuggestsDataLoading()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterList(
            objectType: "measures",
            count: 0);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        Assert.Contains(suggestions, s => s.Contains("No", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetNextStepsAfterList_WithItems_ProvidesContextualActions()
    {
        // Arrange & Act
        var suggestions = DataModelWorkflowGuidance.GetNextStepsAfterList(
            objectType: "measures",
            count: 5);

        // Assert
        Assert.NotNull(suggestions);
        Assert.NotEmpty(suggestions);
        Assert.Contains(suggestions, s => s.Contains("Found", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetErrorRecoverySteps_ProvidesToolSpecificGuidance()
    {
        // Arrange & Act - Test different error types
        var daxSyntaxSteps = DataModelWorkflowGuidance.GetErrorRecoverySteps("DAXSyntax");
        var columnNotFoundSteps = DataModelWorkflowGuidance.GetErrorRecoverySteps("ColumnNotFound");
        var circularDepSteps = DataModelWorkflowGuidance.GetErrorRecoverySteps("CircularDependency");

        // Assert
        Assert.NotEmpty(daxSyntaxSteps);
        Assert.Contains(daxSyntaxSteps, s => s.Contains("DAX", StringComparison.OrdinalIgnoreCase));
        
        Assert.NotEmpty(columnNotFoundSteps);
        Assert.Contains(columnNotFoundSteps, s => s.Contains("column", StringComparison.OrdinalIgnoreCase));
        
        Assert.NotEmpty(circularDepSteps);
        Assert.Contains(circularDepSteps, s => s.Contains("circular", StringComparison.OrdinalIgnoreCase) ||
                                               s.Contains("relationship", StringComparison.OrdinalIgnoreCase));
    }
}
