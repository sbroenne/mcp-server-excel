using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableOlapAllMemberTests
{
    [Theory]
    [InlineData(unchecked((int)0x80070005))]
    [InlineData(unchecked((int)0x800A0CB3))]
    public void MetadataQueryFailure_ReportsUnavailableSelection(int hresult)
    {
#pragma warning disable CA2201 // Inject provider failures without requiring Excel or changing metadata permissions.
        var cube = CreateCube(new FakeConnection(new COMException("Private provider details.", hresult)));
#pragma warning restore CA2201

        var result = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() =>
            PivotTableHelpers.IsAllOlapReportFilterMember(cube, "[Region].[AllMember]", CancellationToken.None));

        Assert.False(result.Available);
        Assert.Contains("connection 'TestConnection'", result.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("TestPivot", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains($"HRESULT 0x{hresult:X8}", result.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("Private provider details", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void NullSchemaRowset_ReportsUnavailableSelection()
    {
        var cube = CreateCube(new FakeConnection(null));

        var result = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() =>
            PivotTableHelpers.IsAllOlapReportFilterMember(cube, "[Region].[AllMember]", CancellationToken.None));

        Assert.False(result.Available);
        Assert.Contains("The OLAP provider returned no result for a schema rowset request.",
            result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void MetadataQueryCancellation_Propagates()
    {
        var failure = new OperationCanceledException("Query cancelled.");
        var cube = CreateCube(new FakeConnection(failure));

        var actual = Assert.Throws<OperationCanceledException>(() =>
            PivotTableSnapshotHelpers.TryReadSnapshotProperty(() =>
                PivotTableHelpers.IsAllOlapReportFilterMember(cube, "[Region].[AllMember]", CancellationToken.None)));

        Assert.Same(failure, actual);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LostExcel_Propagates(bool adoGetterFails)
    {
#pragma warning disable CA2201 // Simulate Excel disconnecting at either metadata boundary.
        var failure = new COMException("Excel disconnected.", unchecked((int)0x80010108));
#pragma warning restore CA2201
        var cube = CreateCube(new FakeConnection(failure), adoGetterFails ? failure : null);

        var actual = Assert.ThrowsAny<Exception>(() =>
            PivotTableSnapshotHelpers.TryReadSnapshotProperty(() =>
                PivotTableHelpers.IsAllOlapReportFilterMember(cube, "[Region].[AllMember]", CancellationToken.None)));

        if (adoGetterFails)
            Assert.Same(failure, actual);
        else
            Assert.Equal(failure.HResult, Assert.IsType<COMException>(actual.InnerException).HResult);
    }

    private static Excel.CubeField CreateCube(FakeConnection connection, Exception? adoFailure = null)
    {
        var workbookConnection = FakeExcelObject.Create<Excel.WorkbookConnection>(name => name switch
        {
            "get_Name" => "TestConnection",
            _ => throw new NotSupportedException(name)
        });
        var cache = FakeExcelObject.Create<Excel.PivotCache>(name => name switch
        {
            "get_CommandText" => "TestCube",
            "get_WorkbookConnection" => workbookConnection,
            "get_ADOConnection" => adoFailure is null
                ? ExcelInterfaceProxy.Create<ADODB._Connection>((member, args) =>
                    member.Name == "OpenSchema"
                        ? connection.OpenSchema(Convert.ToInt32(args![0], System.Globalization.CultureInfo.InvariantCulture),
                            Assert.IsType<object[]>(args[1]))
                        : throw new NotSupportedException(member.Name))
                : throw adoFailure,
            _ => throw new NotSupportedException(name)
        });
        var pivot = FakeExcelObject.Create<Excel.PivotTable>(name => name switch
        {
            "PivotCache" => cache,
            "get_Name" => "TestPivot",
            _ => throw new NotSupportedException(name)
        });
        return FakeExcelObject.Create<Excel.CubeField>(name => name switch
        {
            "get_Parent" => pivot,
            "get_Name" => "[Region]",
            _ => throw new NotSupportedException(name)
        });
    }

    public sealed class FakeConnection(Exception? failure)
    {
        public object? OpenSchema(int schema, object?[] restrictions)
        {
            Assert.Equal(34, schema);
            Assert.Equal(new object?[] { null, null, "TestCube", null, null, "[Region]" }, restrictions);
            if (failure is not null)
                throw failure;
            return null;
        }
    }

    public class FakeExcelObject : DispatchProxy
    {
        private Func<string, object?> _read = null!;

        public static T Create<T>(Func<string, object?> read) where T : class
        {
            var interfaceType = typeof(PivotTableHelpers).Assembly.GetType(typeof(T).FullName!) ?? typeof(T);
            var proxy = DispatchProxy.Create(interfaceType, typeof(FakeExcelObject));
            ((FakeExcelObject)proxy)._read = read;
            return (T)proxy;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            _read(targetMethod!.Name);
    }
}
