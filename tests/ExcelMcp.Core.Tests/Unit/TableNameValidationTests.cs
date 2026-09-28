using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

#pragma warning disable CA1822, CA2201 // Test setters intentionally only throw fabricated failures.

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class TableNameValidationTests
{
    [Fact]
    public void TryAssignTableNameForValidation_WhenExcelRejectsName_ReturnsFalse()
    {
        var table = new RejectingTable();

        bool accepted = TableCommands.TryAssignTableNameForValidation(table, "Invalid Name");

        Assert.False(accepted);
    }

    [Fact]
    public void TryAssignTableNameForValidation_WhenUnexpectedFailureOccurs_Propagates()
    {
        var table = new BrokenTable();

        Assert.Throws<InvalidOperationException>(() =>
            TableCommands.TryAssignTableNameForValidation(table, "Table1"));
    }

    [Theory]
    [InlineData(ResiliencePipelines.RPC_E_SERVERCALL_RETRYLATER)]
    [InlineData(ResiliencePipelines.RPC_E_CALL_REJECTED)]
    [InlineData(ResiliencePipelines.RPC_E_CALL_FAILED)]
    [InlineData(ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE)]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    [InlineData(unchecked((int)0x8007000E))]
    public void TryAssignTableNameForValidation_WhenExcelSessionFails_Propagates(int hResult)
    {
        var table = new ComFailureTable(hResult);

        var exception = Assert.Throws<COMException>(() =>
            TableCommands.TryAssignTableNameForValidation(table, "Table1"));

        Assert.Equal(hResult, exception.HResult);
    }

    // Service calls cannot verify that Core validates the name before using the batch.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_WithMissingTableName_ThrowsBeforeUsingBatch(string? tableName)
    {
        var commands = new TableCommands();
        var exception = Assert.ThrowsAny<ArgumentException>(() =>
            commands.Create(
                null!,
                "Data",
                tableName!,
                "A1:B2",
                true,
                "TableStyleLight1"));

        Assert.Equal("tableName", exception.ParamName);
    }

    public sealed class RejectingTable
    {
        public string Name
        {
            set => throw new COMException("Excel rejected the name.");
        }
    }

    public sealed class BrokenTable
    {
        public string Name
        {
            set => throw new InvalidOperationException("The session failed.");
        }
    }

    public sealed class ComFailureTable(int hResult)
    {
        public string Name
        {
            set => throw new COMException("The Excel session failed.", hResult);
        }
    }
}
