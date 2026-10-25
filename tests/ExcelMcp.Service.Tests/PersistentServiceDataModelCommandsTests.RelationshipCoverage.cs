using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServiceDataModelCommandsTests
{
    [Theory]
    [InlineData("CustomerID", "CustomersTable")]
    [InlineData("ProductID", "ProductsTable")]
    public void ReadRelationship_ReturnsExactSelectedEndpointsAndNativeActiveState(string column, string lookup)
    {
        var before = ReadNativeRelationships();
        Assert.Equal(2, before.Length);
        var result = RequireSuccess(_dataModelCommands.ReadRelationship(
            _fixture.BatchToken, "SalesTable", column, lookup, column));
        var relationship = Assert.IsType<DataModelRelationshipInfo>(result.Relationship);
        Assert.Equal(("SalesTable", column, lookup, column, true), RelationshipKey(relationship));
        Assert.Contains(RelationshipKey(relationship), before);
        Assert.Equal(before, ReadNativeRelationships());
    }

    [Fact]
    public void UpdateRelationship_TogglesActiveStateWithoutChangingEndpointsOrOtherRelationship()
    {
        var batch = _fixture.BatchToken;
        var before = ReadNativeRelationships();
        Assert.Equal(new[]
        {
            ("SalesTable", "CustomerID", "CustomersTable", "CustomerID", true),
            ("SalesTable", "ProductID", "ProductsTable", "ProductID", true)
        }, before);
        PersistentServiceCleanupFailures.Run(() =>
        {
            RequireSuccess(_dataModelCommands.UpdateRelationship(
                batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: false));
            Assert.Equal(new[]
            {
                ("SalesTable", "CustomerID", "CustomersTable", "CustomerID", false),
                ("SalesTable", "ProductID", "ProductsTable", "ProductID", true)
            }, ReadNativeRelationships());
            var inactive = RequireSuccess(_dataModelCommands.ReadRelationship(
                batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID"));
            Assert.Equal(("SalesTable", "CustomerID", "CustomersTable", "CustomerID", false),
                RelationshipKey(Assert.IsType<DataModelRelationshipInfo>(inactive.Relationship)));
            RequireSuccess(_dataModelCommands.UpdateRelationship(
                batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true));
            Assert.Equal(before, ReadNativeRelationships());
            var active = RequireSuccess(_dataModelCommands.ReadRelationship(
                batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID"));
            Assert.True(Assert.IsType<DataModelRelationshipInfo>(active.Relationship).IsActive);
        }, () => RequireSuccess(_dataModelCommands.UpdateRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true)));
    }

    private static (string, string, string, string, bool) RelationshipKey(DataModelRelationshipInfo relationship) =>
        (relationship.FromTable, relationship.FromColumn, relationship.ToTable, relationship.ToColumn, relationship.IsActive);

    private (string, string, string, string, bool)[] ReadNativeRelationships() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Model? model = null;
            Excel.ModelRelationships? relationships = null;
            Excel.ModelRelationship? relationship = null;
            Excel.ModelTableColumn? foreignColumn = null;
            Excel.ModelTableColumn? primaryColumn = null;
            Excel.ModelTable? foreignTable = null;
            Excel.ModelTable? primaryTable = null;
            var result = new List<(string, string, string, string, bool)>();
            try
            {
                model = context.Book.Model;
                relationships = model.ModelRelationships;
                for (int index = 1; index <= relationships.Count; index++)
                {
                    try
                    {
                        relationship = relationships.Item(index);
                        foreignColumn = relationship.ForeignKeyColumn;
                        primaryColumn = relationship.PrimaryKeyColumn;
                        foreignTable = (Excel.ModelTable)foreignColumn.Parent;
                        primaryTable = (Excel.ModelTable)primaryColumn.Parent;
                        result.Add((foreignTable.Name, foreignColumn.Name,
                            primaryTable.Name, primaryColumn.Name, relationship.Active));
                    }
                    finally
                    {
                        ComUtilities.Release(ref primaryTable);
                        ComUtilities.Release(ref foreignTable);
                        ComUtilities.Release(ref primaryColumn);
                        ComUtilities.Release(ref foreignColumn);
                        ComUtilities.Release(ref relationship);
                    }
                }
                return result.OrderBy(item => item.Item2, StringComparer.Ordinal).ToArray();
            }
            finally
            {
                ComUtilities.Release(ref relationships);
                ComUtilities.Release(ref model);
            }
        });
}
