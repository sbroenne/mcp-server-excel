using System.Reflection;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public sealed class ManualToolParameterContractTests
{
    [Theory]
    [InlineData("file", "file_path", "open", true)]
    [InlineData("file", "file_path", "create", true)]
    [InlineData("file", "workbook_session_id", "close", true)]
    [InlineData("file", "save", "close", false)]
    [InlineData("file", "show", "open", false)]
    [InlineData("file", "show", "create", false)]
    [InlineData("file", "timeout_seconds", "open", false)]
    [InlineData("file", "timeout_seconds", "create", false)]
    [InlineData("file_read", "file_path", "test", true)]
    [InlineData("file_read", "timeout_seconds", "test", false)]
    public void FileToolParameters_DeclareActionApplicabilityAndRequiredness(
        string toolName,
        string parameterName,
        string action,
        bool required)
    {
        var method = typeof(ExcelFileTool).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Single(method => method.GetCustomAttribute<McpServerToolAttribute>()?.Name == toolName);
        var parameter = method.GetParameters().Single(parameter => parameter.Name == parameterName);
        var contract = Assert.Single(
            parameter.GetCustomAttributesData(),
            attribute => attribute.AttributeType.Name == "McpActionParameterAttribute" &&
                Equals(attribute.ConstructorArguments[0].Value, action));

        Assert.Equal(action, contract.ConstructorArguments[0].Value);
        Assert.Equal(
            required,
            contract.NamedArguments.SingleOrDefault(argument => argument.MemberName == "Required")
                .TypedValue.Value as bool? ?? false);
    }
}
