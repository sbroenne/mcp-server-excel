using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class TableNameRollbackTests
{
    [Fact]
    public void SetCreatedTableNameOrRollback_RangeTableNameRejected_UnlistsTable()
    {
        var table = new RejectingTable();

        Assert.Throws<COMException>(() =>
            TableCommands.SetCreatedTableNameOrRollback(
                table,
                "Rejected",
                preserveSourceRange: true));

        Assert.Equal("Rejected", table.AttemptedName);
        Assert.True(table.Unlisted);
        Assert.False(table.Deleted);
    }

    [Fact]
    public void SetCreatedTableNameOrRollback_DaxTableNameRejected_DeletesTableAndConnection()
    {
        var table = new RejectingTable();
        var connection = new DeletableConnection();

        Assert.Throws<COMException>(() =>
            TableCommands.SetCreatedTableNameOrRollback(
                table,
                "Rejected",
                connection));

        Assert.Equal("Rejected", table.AttemptedName);
        Assert.True(table.Deleted);
        Assert.False(table.Unlisted);
        Assert.True(connection.Deleted);
    }

    public sealed class RejectingTable
    {
        public string Name
        {
            set
            {
                AttemptedName = value;
                throw Marshal.GetExceptionForHR(unchecked((int)0x800A03EC))!;
            }
        }

        public string? AttemptedName { get; private set; }
        public bool Deleted { get; private set; }
        public bool Unlisted { get; private set; }

        public void Delete() => Deleted = true;
        public void Unlist() => Unlisted = true;
    }

    public sealed class DeletableConnection
    {
        public bool Deleted { get; private set; }

        public void Delete() => Deleted = true;
    }
}
