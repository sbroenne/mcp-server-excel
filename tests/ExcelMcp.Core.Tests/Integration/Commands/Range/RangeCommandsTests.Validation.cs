// Copyright (c) Stefan Broenne. All rights reserved.

using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    public void ValidateRange_WithInputMessage_ReturnsSuccess()
    {
        // Arrange & Act - First write list values to worksheet (required for dropdown)
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(
            batch,
            sheetName,
            "B1:B3",
            new List<List<object?>>
            {
                new() { "Option1" },
                new() { "Option2" },
                new() { "Option3" }
            });

        // Apply validation referencing the range (creates dropdown)
        _commands.ValidateRange(
            batch,
            sheetName,
            "A1",
            validationType: "list",
            validationOperator: null,
            formula1: "=$B$1:$B$3",  // Reference to worksheet range creates dropdown
            formula2: null,
            showInputMessage: true,
            inputTitle: "My Input Title",
            inputMessage: "My helpful input message",
            showErrorAlert: true,
            errorStyle: "stop",
            errorTitle: "My Error Title",
            errorMessage: "My error message",
            ignoreBlank: true,
            showDropdown: true);
        // void method throws on failure, succeeds silently

        // Verify validation is retrieved correctly (same batch)
        var getResult = _commands.GetValidation(batch, sheetName, "A1");

        // Assert - Validation retrieved successfully
        Assert.True(getResult.Success, $"Get validation failed: {getResult.ErrorMessage}");
        Assert.True(getResult.HasValidation, "Range should have validation");

        // Assert - Validation type and formula are correct
        Assert.Equal("list", getResult.ValidationType);
        Assert.Equal("=$B$1:$B$3", getResult.Formula1);

        // Assert - Input message properties are returned
        Assert.Equal("My Input Title", getResult.InputTitle);
        Assert.Equal("My helpful input message", getResult.InputMessage);

        // Assert - Error message properties are returned (these work according to the bug report)
        Assert.Equal("My Error Title", getResult.ErrorTitle);
        Assert.Equal("My error message", getResult.ValidationErrorMessage);
    }

    [Fact]
    public void GetValidation_WithInputMessage_ReturnsInputTitleAndMessage()
    {
        // Arrange & Act - First write list values to worksheet (required for dropdown)
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(
            batch,
            sheetName,
            "B1:B3",
            new List<List<object?>>
            {
                new() { "Option1" },
                new() { "Option2" },
                new() { "Option3" }
            });

        // Apply validation using the ValidateRangeAsync API
        _commands.ValidateRange(
            batch,
            sheetName,
            "A1",
            validationType: "list",
            validationOperator: null,  // Not used for list type
            formula1: "=$B$1:$B$3",  // Reference to worksheet range creates dropdown
            formula2: null,
            showInputMessage: true,
            inputTitle: "My Input Title",
            inputMessage: "My helpful input message",
            showErrorAlert: true,
            errorStyle: "stop",
            errorTitle: "My Error Title",
            errorMessage: "My error message",
            ignoreBlank: true,
            showDropdown: true);
        // void method throws on failure, succeeds silently

        // Act - Get validation to verify InputTitle/InputMessage are returned
        var result = _commands.GetValidation(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"Get validation failed: {result.ErrorMessage}");
        Assert.True(result.HasValidation, "Range should have validation");

        // Assert - Validation type and formula create dropdown with 3 values
        Assert.Equal("list", result.ValidationType);
        Assert.Equal("=$B$1:$B$3", result.Formula1);

        // CRITICAL: These assertions test the bug fix
        Assert.NotEmpty(result.InputTitle ?? string.Empty);
        Assert.Equal("My Input Title", result.InputTitle);
        Assert.NotEmpty(result.InputMessage ?? string.Empty);
        Assert.Equal("My helpful input message", result.InputMessage);

        // These should work (error properties worked before the fix)
        Assert.Equal("My Error Title", result.ErrorTitle);
        Assert.Equal("My error message", result.ValidationErrorMessage);
    }
}
