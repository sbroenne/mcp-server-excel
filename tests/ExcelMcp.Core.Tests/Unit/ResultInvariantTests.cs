using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Results")]
[Trait("RequiresExcel", "false")]
public sealed class ResultInvariantTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ErrorAlwaysPreventsSuccessRegardlessOfAssignmentOrder(bool errorFirst)
    {
        var result = new OperationResult();
        if (errorFirst)
        {
            result.ErrorMessage = "Native operation failed.";
            result.Success = true;
        }
        else
        {
            result.Success = true;
            result.ErrorMessage = "Native operation failed.";
        }

        Assert.False(result.Success);
        Assert.Equal("Native operation failed.", result.ErrorMessage);
        Assert.False(JsonSerializer.SerializeToElement(result).GetProperty("Success").GetBoolean());
    }

    [Theory]
    [InlineData("""{"Success":true,"ErrorMessage":"Native operation failed."}""")]
    [InlineData("""{"ErrorMessage":"Native operation failed.","Success":true}""")]
    public void DeserializationNeverPublishesSuccessWithAnError(string json)
    {
        var result = JsonSerializer.Deserialize<OperationResult>(json)!;

        Assert.False(result.Success);
        Assert.Equal("Native operation failed.", result.ErrorMessage);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void EmptyErrorDoesNotPromoteFailureToSuccess(string? error)
    {
        var result = new OperationResult { Success = false, ErrorMessage = error };

        Assert.False(result.Success);
        result.Success = true;
        Assert.True(result.Success);
    }
}
