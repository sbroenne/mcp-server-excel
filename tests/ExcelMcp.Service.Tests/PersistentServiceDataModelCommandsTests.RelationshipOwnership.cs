using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServiceDataModelCommandsTests
{
    [Theory]
    [InlineData("update")]
    [InlineData("duplicate")]
    [InlineData("missing")]
    [InlineData("inspection-failure")]
    public void RelationshipLookup_ReleasesColumnsAndTablesOnEveryExit(string operation)
    {
        var batch = _fixture.BatchToken;
        var before = RequireSuccess(_dataModelCommands.ListRelationships(batch)).Relationships;
        var acquired = new List<object>();
        DataModelCommands.RelationshipReferenceAcquiredForTests = reference =>
        {
            acquired.Add(reference);
            if (operation == "inspection-failure" && acquired.Count == 2)
                throw new InvalidOperationException("Injected relationship inspection failure.");
        };
        PersistentServiceCleanupFailures.Run(() =>
        {
            if (operation == "update")
            {
                RequireSuccess(_dataModelCommands.UpdateRelationship(
                    batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: false));
            }
            else
            {
                var failure = Assert.Throws<InvalidOperationException>(() =>
                {
                    if (operation == "duplicate")
                    {
                        _dataModelCommands.CreateRelationship(
                            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");
                    }
                    else
                    {
                        _dataModelCommands.UpdateRelationship(
                            batch, "SalesTable", operation == "missing" ? "SalesID" : "CustomerID",
                            "CustomersTable", "CustomerID", active: false);
                    }
                });
                Assert.Contains(operation switch
                {
                    "duplicate" => "already exists",
                    "inspection-failure" => "Injected relationship inspection failure.",
                    _ => "not found"
                }, failure.Message, StringComparison.OrdinalIgnoreCase);
            }

            Assert.NotEmpty(acquired);
            _fixture.ExecuteRawVerification((_, _) =>
            {
                foreach (var reference in acquired)
                {
                    Assert.Throws<InvalidComObjectException>(() =>
                    {
                        if (reference is Excel.ModelTableColumn column)
                            _ = column.Name;
                        else
                            _ = ((Excel.ModelTable)reference).Name;
                    });
                }
            });

            DataModelCommands.RelationshipReferenceAcquiredForTests = null;
            var after = RequireSuccess(_dataModelCommands.ListRelationships(batch)).Relationships;
            Assert.Equal(before.Count, after.Count);
            foreach (var expected in before)
            {
                var actual = Assert.Single(after, relationship =>
                    relationship.FromTable == expected.FromTable &&
                    relationship.FromColumn == expected.FromColumn &&
                    relationship.ToTable == expected.ToTable &&
                    relationship.ToColumn == expected.ToColumn);
                Assert.Equal(operation == "update" && expected.FromTable == "SalesTable" &&
                    expected.FromColumn == "CustomerID" && expected.ToTable == "CustomersTable"
                        ? false : expected.IsActive, actual.IsActive);
            }
        },
        () =>
        {
            DataModelCommands.RelationshipReferenceAcquiredForTests = null;
            _fixture.ExecuteRawVerification((_, _) =>
            {
                foreach (var reference in acquired)
                {
                    object? remaining = reference;
                    ComUtilities.Release(ref remaining);
                }
            });
        },
        () =>
        {
            if (operation == "update")
                RequireSuccess(_dataModelCommands.UpdateRelationship(
                    batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID",
                    active: Assert.Single(before, relationship =>
                        relationship.FromTable == "SalesTable" &&
                        relationship.FromColumn == "CustomerID" &&
                        relationship.ToTable == "CustomersTable").IsActive));
        });
    }
}
