using System.Reflection;
using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for Excel-independent table name validation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class TableNameValidationTests
{
    private static readonly MethodInfo _validateTableName = typeof(TableCommands).GetMethod(
        "ValidateTableName",
        BindingFlags.NonPublic | BindingFlags.Static)!;

    [Theory]
    [InlineData("表1")]
    [InlineData("テーブル1")]
    [InlineData("표1")]
    [InlineData("Tâblé1")]
    public void ValidateTableName_WithLocalizedExcelTableName_DoesNotThrow(string tableName)
    {
        InvokeValidateTableName(tableName);
    }

    [Fact]
    public void ValidateTableName_WithLeadingDigit_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => InvokeValidateTableName("1Table"));
    }

    private static void InvokeValidateTableName(string tableName)
    {
        try
        {
            _validateTableName.Invoke(null, [tableName]);
        }
        catch (TargetInvocationException ex) when (ex.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
        }
    }
}
