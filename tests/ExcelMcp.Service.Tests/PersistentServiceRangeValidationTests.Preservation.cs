using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeValidationTests
{
    [Theory]
    [InlineData("unsupported", "between", "stop")]
    [InlineData("whole", "unsupported", "stop")]
    [InlineData("whole", "between", "unsupported")]
    public void ValidateRange_InvalidOptions_PreserveExistingRule(
        string validationType, string validationOperator, string errorStyle)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.ValidateRange(
            batch, sheet, "A1", "whole", "between", "1", "10",
            true, "Original title", "Original message",
            true, "warning", "Original error", "Original error message",
            false, null));
        var before = _commands.GetValidation(batch, sheet, "A1");
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.HasValidation);

        Assert.Throws<ArgumentException>(() => _commands.ValidateRange(
            batch, sheet, "A1", validationType, validationOperator, "2", "8",
            true, "Replacement title", "Replacement message",
            true, errorStyle, "Replacement error", "Replacement error message",
            true, null));

        var after = _commands.GetValidation(batch, sheet, "A1");
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(before), JsonSerializer.Serialize(after));
    }
}
