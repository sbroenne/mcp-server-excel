// Copyright (c) Stefan Broenne. All rights reserved.

using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeValidationTests
{
    [Fact]
    public void ValidateRange_WithInputMessage_ReturnsSuccess()
    {
        // Arrange & Act - First write list values to worksheet (required for dropdown)
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var written = _commands.SetValues(
            batch,
            sheetName,
            "B1:B3",
            new List<List<object?>>
            {
                new() { "Option1" },
                new() { "Option2" },
                new() { "Option3" }
            });
        Assert.True(written.Success, written.ErrorMessage);

        // Apply validation referencing the range (creates dropdown)
        var applied = _commands.ValidateRange(
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
        Assert.True(applied.Success, applied.ErrorMessage);

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
        Assert.True(getResult.ShowInputMessage);
        Assert.True(getResult.ShowErrorAlert);
        Assert.True(getResult.IgnoreBlank);
        Assert.Equal("stop", getResult.ErrorStyle);
        Assert.True(ReadDropdown(sheetName, "A1"));
    }

    [Fact]
    public void GetValidation_WithInputMessage_ReturnsInputTitleAndMessage()
    {
        // Arrange & Act - First write list values to worksheet (required for dropdown)
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var written = _commands.SetValues(
            batch,
            sheetName,
            "B1:B3",
            new List<List<object?>>
            {
                new() { "Option1" },
                new() { "Option2" },
                new() { "Option3" }
            });
        Assert.True(written.Success, written.ErrorMessage);

        // Apply validation using the ValidateRangeAsync API
        var applied = _commands.ValidateRange(
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
        Assert.True(applied.Success, applied.ErrorMessage);

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
        Assert.True(result.ShowInputMessage);
        Assert.True(result.ShowErrorAlert);
        Assert.True(result.IgnoreBlank);
        Assert.Equal("stop", result.ErrorStyle);
        Assert.True(ReadDropdown(sheetName, "A1"));
    }

    [Theory]
    [InlineData(true, "stop")]
    [InlineData(false, "stop")]
    [InlineData(true, "warning")]
    [InlineData(false, "warning")]
    [InlineData(true, "information")]
    [InlineData(false, "information")]
    public void ValidateRange_ExplicitFlagsAndErrorStyle_StoresRequestedSettings(bool enabled, string style)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "B1:B3",
            [["First"], ["Second"], ["Third"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["First"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [["Untouched"]]).Success);
        Assert.True(_commands.ValidateRange(batch, sheetName, "A1", "list", "between",
            "=$B$1:$B$3", null, true, "Original input", "Original message",
            true, "stop", "Original error", "Original error message", true, true).Success);

        var applied = _commands.ValidateRange(batch, sheetName, "A1", "list", "between",
            "=$B$1:$B$3", null, enabled, null, null,
            enabled, style, null, null, enabled, enabled);

        Assert.True(applied.Success, applied.ErrorMessage);
        var result = _commands.GetValidation(batch, sheetName, "A1");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.HasValidation);
        Assert.Equal("list", result.ValidationType);
        Assert.Equal("=$B$1:$B$3", result.Formula1);
        Assert.Equal(enabled, result.ShowInputMessage);
        Assert.Equal(enabled, result.ShowErrorAlert);
        Assert.Equal(enabled, result.IgnoreBlank);
        Assert.Equal(style, result.ErrorStyle);
        Assert.Equal(enabled, ReadDropdown(sheetName, "A1"));
        var cells = _commands.GetValues(batch, sheetName, "A1:C1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["First", "First", "Untouched"], Assert.Single(cells.Values));
    }

    [Theory]
    [InlineData("unsupported", "between", "stop", "Invalid validation type")]
    [InlineData("list", "unsupported", "stop", "Invalid validation operator")]
    [InlineData("list", "between", "unsupported", "Invalid error style")]
    public void ValidateRange_InvalidOption_PreservesExistingRuleAndCells(
        string validationType, string validationOperator, string errorStyle, string expectedError)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "B1:B3",
            [["First"], ["Second"], ["Third"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["First"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [["Untouched"]]).Success);
        Assert.True(_commands.ValidateRange(batch, sheetName, "A1", "list", "between",
            "=$B$1:$B$3", null, true, "Input title", "Input message",
            true, "warning", "Error title", "Error message", false, false).Success);
        var before = _commands.GetValidation(batch, sheetName, "A1");
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.HasValidation);
        Assert.Equal("list", before.ValidationType);
        Assert.Equal("=$B$1:$B$3", before.Formula1);
        Assert.Equal("warning", before.ErrorStyle);
        Assert.Equal("Input title", before.InputTitle);
        Assert.Equal("Input message", before.InputMessage);
        Assert.Equal("Error title", before.ErrorTitle);
        Assert.Equal("Error message", before.ValidationErrorMessage);
        Assert.True(before.ShowInputMessage);
        Assert.True(before.ShowErrorAlert);
        Assert.False(before.IgnoreBlank);
        Assert.False(ReadDropdown(sheetName, "A1"));
        var cellsBefore = _commands.GetValues(batch, sheetName, "A1:C3");
        Assert.True(cellsBefore.Success, cellsBefore.ErrorMessage);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.ValidateRange(batch, sheetName, "A1", validationType, validationOperator,
                "=$B$1:$B$3", null, false, null, null,
                false, errorStyle, null, null, true, true));

        Assert.Contains(expectedError, exception.Message, StringComparison.Ordinal);
        var after = _commands.GetValidation(batch, sheetName, "A1");
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(before), JsonSerializer.Serialize(after));
        Assert.False(ReadDropdown(sheetName, "A1"));
        var cellsAfter = _commands.GetValues(batch, sheetName, "A1:C3");
        Assert.True(cellsAfter.Success, cellsAfter.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(cellsBefore.Values), JsonSerializer.Serialize(cellsAfter.Values));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(true)]
    [InlineData(false)]
    public void ValidateRange_DefaultOptions_StoresErrorAlertTextWhenEnabled(bool? showErrorAlert)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "B1:B3",
            [["First"], ["Second"], ["Third"]]).Success);
        Assert.True(_commands.ValidateRange(batch, sheetName, "A1", "list", null,
            "=$B$1:$B$3", null, true, "Original input", "Original message",
            false, "warning", null, null, false, false).Success);

        var applied = _commands.ValidateRange(batch, sheetName, "A1", "list", null,
            "=$B$1:$B$3", null, null, null, null,
            showErrorAlert, null, "Custom error title", "Custom error message", null, null);

        Assert.True(applied.Success, applied.ErrorMessage);
        var result = _commands.GetValidation(batch, sheetName, "A1");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.HasValidation);
        Assert.Equal("list", result.ValidationType);
        Assert.Equal("=$B$1:$B$3", result.Formula1);
        Assert.False(result.ShowInputMessage);
        Assert.Equal(showErrorAlert ?? true, result.ShowErrorAlert);
        Assert.Equal(showErrorAlert is false ? "" : "Custom error title", result.ErrorTitle);
        Assert.Equal(showErrorAlert is false ? "" : "Custom error message", result.ValidationErrorMessage);
        Assert.True(result.IgnoreBlank);
        Assert.Equal("stop", result.ErrorStyle);
        Assert.True(ReadDropdown(sheetName, "A1"));
    }

    private bool ReadDropdown(string sheetName, string address) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                range = sheet.Range[address];
                validation = range.Validation;
                return Convert.ToBoolean(validation.InCellDropdown, CultureInfo.InvariantCulture);
            }
            finally
            {
                ComUtilities.Release(ref validation);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
}
