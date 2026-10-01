using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPythonInExcelArgumentsTests
{
    [Fact]
    public void PrepareSetFormula_PreservesCodeAndAddsExplicitDefaultReturnType()
    {
        var arguments = new JsonObject
        {
            ["sheetName"] = "Data",
            ["rangeAddress"] = "D1",
            ["code"] = "print(\"quoted\")"
        };

        MacPythonInExcelArguments.Prepare("set-formula", arguments, TimeSpan.FromSeconds(60));

        Assert.Equal("print(\"quoted\")", arguments["code"]?.GetValue<string>());
        Assert.Equal(0, arguments["returnType"]?.GetValue<int>());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public void PrepareSetFormula_RejectsMissingOrBlankCode(string? code)
    {
        var arguments = RequiredArguments();
        if (code is not null)
        {
            arguments["code"] = code;
        }

        var error = Assert.Throws<ArgumentException>(() =>
            MacPythonInExcelArguments.Prepare("set-formula", arguments, TimeSpan.FromSeconds(60)));

        Assert.Contains("Python code must not be empty.", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareGetResult_AddsContractDefaultWaitAndOperationDeadline()
    {
        var arguments = RequiredArguments();

        MacPythonInExcelArguments.Prepare("get-result", arguments, TimeSpan.FromSeconds(60));

        Assert.Equal(30, arguments["maxWaitSeconds"]?.GetValue<int>());
        Assert.Equal(60, arguments["operationTimeoutSeconds"]?.GetValue<double>());
    }

    [Theory]
    [InlineData(0, 60)]
    [InlineData(60, 60)]
    [InlineData(61, 60)]
    public void PrepareGetResult_RejectsWaitOutsideOperationDeadline(int wait, int operationTimeout)
    {
        var arguments = RequiredArguments();
        arguments["maxWaitSeconds"] = wait;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            MacPythonInExcelArguments.Prepare(
                "get-result",
                arguments,
                TimeSpan.FromSeconds(operationTimeout)));
    }

    [Theory]
    [InlineData("sheetName")]
    [InlineData("rangeAddress")]
    public void Prepare_RejectsMissingRequiredLocation(string property)
    {
        var arguments = RequiredArguments();
        arguments.Remove(property);
        arguments["code"] = "1 + 1";

        var error = Assert.Throws<ArgumentException>(() =>
            MacPythonInExcelArguments.Prepare("set-formula", arguments, TimeSpan.FromSeconds(60)));

        Assert.Contains($"{property} is required.", error.Message, StringComparison.Ordinal);
    }

    private static JsonObject RequiredArguments() =>
        new()
        {
            ["sheetName"] = "Data",
            ["rangeAddress"] = "D1"
        };
}
