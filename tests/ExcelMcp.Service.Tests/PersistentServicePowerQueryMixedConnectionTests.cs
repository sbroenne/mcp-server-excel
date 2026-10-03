using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryMixedConnectionTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string ModelCode =
        "let Source = #table(type table [Guard = Int64.Type], {{71}, {89}}) in Source";
    private const string OriginalCode = "let Source = #table({\"A\"}, {{1}, {3}}) in Source";
    private const string UpdatedCode = "let Source = #table({\"A\", \"B\"}, {{7, 11}, {23, 31}}) in Source";
    private readonly IPowerQueryCommands _queries = fixture.CreateCommands<IPowerQueryCommands>();

    [Fact]
    public void Update_WithDataModelConnection_DoesNotThrowCOMException()
    {
        var guard = CreateModelGuard();
        var name = CreateQuery(PowerQueryLoadMode.ConnectionOnly);
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, UpdatedCode));
        AssertStored(name, UpdatedCode, PowerQueryLoadMode.ConnectionOnly);
        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, UpdatedCode));
        Assert.Equal(["A", "B"], result.Columns);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        PowerQueryStateAssertions.AssertRows([[7, 11], [23, 31]], result.Rows);
        AssertModelGuard(guard);
        AssertStored(name, UpdatedCode, PowerQueryLoadMode.ConnectionOnly);
    }

    [Fact]
    public void View_WithDataModelConnection_DoesNotThrowCOMException()
    {
        var guard = CreateModelGuard();
        var name = CreateQuery(PowerQueryLoadMode.ConnectionOnly);
        AssertStored(name, OriginalCode, PowerQueryLoadMode.ConnectionOnly);
        AssertModelGuard(guard);
    }

    [Fact]
    public void List_WithDataModelConnection_Succeeds()
    {
        var guard = CreateModelGuard();
        var name = CreateQuery(PowerQueryLoadMode.ConnectionOnly);
        var result = RequireSuccess(_queries.List(_fixture.BatchToken));
        Assert.Equal(new[] { guard, name }.Order(StringComparer.Ordinal),
            result.Queries.Select(query => query.Name).Order(StringComparer.Ordinal));
        AssertStored(name, OriginalCode, PowerQueryLoadMode.ConnectionOnly);
        AssertModelGuard(guard);
    }

    [Fact]
    public void Delete_WithDataModelConnection_Succeeds()
    {
        var guard = CreateModelGuard();
        var name = CreateQuery(PowerQueryLoadMode.ConnectionOnly);
        RequireSuccess(_queries.Delete(_fixture.BatchToken, name));
        _fixture.ForgetPowerQuery(name);
        PowerQueryStateAssertions.AssertRemoved(_fixture, name);
        AssertModelGuard(guard);
    }

    [Fact]
    public void Update_WorksheetQuery_WithDataModelConnection_Succeeds()
    {
        var guard = CreateModelGuard();
        var name = CreateQuery(PowerQueryLoadMode.LoadToTable);
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, UpdatedCode));
        AssertStored(name, UpdatedCode, PowerQueryLoadMode.LoadToTable);
        AssertModelGuard(guard);
    }

    [Fact]
    public void DataModelLoad_CreatesNonOledbConnection()
    {
        var guard = CreateModelGuard();
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Model? model = null;
            Excel.WorkbookConnection? connection = null;
            try
            {
                model = context.Book.Model;
                connection = model.DataModelConnection;
                Assert.Equal(Excel.XlConnectionType.xlConnectionTypeMODEL, connection.Type);
                Assert.NotEmpty(connection.Name);
            }
            finally
            {
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref model);
            }
        });
        AssertModelGuard(guard);
    }

    private string CreateModelGuard()
    {
        var name = UniqueName("Model");
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, ModelCode,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(name);
        AssertModelGuard(name);
        return name;
    }

    private string CreateQuery(PowerQueryLoadMode mode)
    {
        var name = UniqueName("Target");
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, OriginalCode, mode,
            mode == PowerQueryLoadMode.LoadToTable ? name : null));
        _fixture.RegisterPowerQueryForCleanup(name);
        if (mode == PowerQueryLoadMode.LoadToTable) { _fixture.RegisterSheetForCleanup(name); }
        AssertStored(name, OriginalCode, mode);
        return name;
    }

    private void AssertStored(string name, string code, PowerQueryLoadMode mode) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, code, mode,
            mode == PowerQueryLoadMode.LoadToTable ? name : null,
            code == OriginalCode ? ["A"] : ["A", "B"],
            code == OriginalCode ? [[1], [3]] : [[7, 11], [23, 31]]);

    private void AssertModelGuard(string name) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, ModelCode,
            PowerQueryLoadMode.LoadToDataModel, null, ["Guard"], [[71], [89]]);

    private static string UniqueName(string prefix) => $"PQ_{prefix}_{Guid.NewGuid():N}"[..(prefix.Length + 12)];
}
