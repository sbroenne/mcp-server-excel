using Sbroenne.ExcelMcp.Core.Models.Validation;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit.Validation;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ActionDefinitionsTests
{
    [Fact]
    public void PowerQuery_List_HasCorrectProperties()
    {
        // Act
        var action = ActionDefinitions.PowerQuery.List;

        // Assert
        Assert.Equal("PowerQuery", action.Domain);
        Assert.Equal("list", action.Name);
        Assert.NotNull(action.Parameters);
        Assert.Single(action.Parameters); // ExcelPath only
        Assert.Equal("excelPath", action.Parameters[0].Name);
        Assert.True(action.Parameters[0].Required);
    }

    [Fact]
    public void PowerQuery_View_HasCorrectParameters()
    {
        // Act
        var action = ActionDefinitions.PowerQuery.View;

        // Assert
        Assert.Equal("PowerQuery", action.Domain);
        Assert.Equal("view", action.Name);
        Assert.Equal(2, action.Parameters.Length);
        Assert.Contains(action.Parameters, p => p.Name == "excelPath");
        Assert.Contains(action.Parameters, p => p.Name == "queryName");
    }

    [Fact]
    public void Parameter_Create_HasCorrectParameters()
    {
        // Act
        var action = ActionDefinitions.Parameter.Create;

        // Assert
        Assert.Equal("Parameter", action.Domain);
        Assert.Equal("create", action.Name);
        Assert.Equal(3, action.Parameters.Length);
        Assert.Contains(action.Parameters, p => p.Name == "excelPath");
        Assert.Contains(action.Parameters, p => p.Name == "parameterName");
        Assert.Contains(action.Parameters, p => p.Name == "reference");
    }

    [Fact]
    public void Table_List_HasCorrectProperties()
    {
        // Act
        var action = ActionDefinitions.Table.List;

        // Assert
        Assert.Equal("Table", action.Domain);
        Assert.Equal("list", action.Name);
        Assert.Single(action.Parameters);
        Assert.Equal("excelPath", action.Parameters[0].Name);
    }

    [Fact]
    public void FindAction_WithValidDomainAndName_ReturnsAction()
    {
        // Act
        var action = ActionDefinitions.FindAction("PowerQuery", "list");

        // Assert
        Assert.NotNull(action);
        Assert.Equal("PowerQuery", action.Domain);
        Assert.Equal("list", action.Name);
    }

    [Fact]
    public void FindAction_WithInvalidDomain_ReturnsNull()
    {
        // Act
        var action = ActionDefinitions.FindAction("InvalidDomain", "list");

        // Assert
        Assert.Null(action);
    }

    [Fact]
    public void FindAction_WithInvalidName_ReturnsNull()
    {
        // Act
        var action = ActionDefinitions.FindAction("PowerQuery", "invalidAction");

        // Assert
        Assert.Null(action);
    }

    [Fact]
    public void GetAllActions_ReturnsAllDefinedActions()
    {
        // Act
        var actions = ActionDefinitions.GetAllActions().ToList();

        // Assert
        Assert.NotEmpty(actions);
        // Should have at least PowerQuery (12) + Parameter (5) + Table (5) = 22 actions
        Assert.True(actions.Count >= 22);
        
        // Verify some exist
        Assert.Contains(actions, a => a.Domain == "PowerQuery" && a.Name == "list");
        Assert.Contains(actions, a => a.Domain == "Parameter" && a.Name == "create");
        Assert.Contains(actions, a => a.Domain == "Table" && a.Name == "rename");
    }

    [Fact]
    public void ActionDefinition_ValidateParameters_WithValidValues_ReturnsSuccess()
    {
        // Arrange
        var action = ActionDefinitions.PowerQuery.View;
        var parameters = new Dictionary<string, object?>
        {
            { "excelPath", "test.xlsx" },
            { "queryName", "Sales" }
        };

        // Act
        var result = action.ValidateParameters(parameters);

        // Assert
        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public void ActionDefinition_ValidateParameters_WithMissingRequired_ReturnsFailure()
    {
        // Arrange
        var action = ActionDefinitions.PowerQuery.View;
        var parameters = new Dictionary<string, object?>
        {
            { "excelPath", "test.xlsx" }
            // Missing queryName
        };

        // Act
        var result = action.ValidateParameters(parameters);

        // Assert
        Assert.False(result.IsValid);
        Assert.Contains("required", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("queryName", result.ParameterName);
    }

    [Fact]
    public void ActionDefinition_ValidateParameters_WithInvalidFileExtension_ReturnsFailure()
    {
        // Arrange
        var action = ActionDefinitions.PowerQuery.List;
        var parameters = new Dictionary<string, object?>
        {
            { "excelPath", "test.txt" } // Invalid extension
        };

        // Act
        var result = action.ValidateParameters(parameters);

        // Assert
        Assert.False(result.IsValid);
        Assert.Contains("must have extension", result.ErrorMessage!);
    }

    [Fact]
    public void ActionDefinition_GetRequiredParameterNames_ReturnsCorrectNames()
    {
        // Arrange
        var action = ActionDefinitions.PowerQuery.Import;

        // Act
        var requiredParams = action.GetRequiredParameterNames();

        // Assert
        Assert.Contains("excelPath", requiredParams);
        Assert.Contains("queryName", requiredParams);
        Assert.Contains("sourcePath", requiredParams);
    }

    [Fact]
    public void ActionDefinition_GetOptionalParameterNames_ReturnsCorrectNames()
    {
        // Arrange
        var action = ActionDefinitions.PowerQuery.Import;

        // Act
        var optionalParams = action.GetOptionalParameterNames();

        // Assert - Import has only privacyLevel as optional
        Assert.Contains("privacyLevel", optionalParams);
    }
}
