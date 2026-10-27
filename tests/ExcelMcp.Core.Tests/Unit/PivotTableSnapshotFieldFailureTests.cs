using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

// Inject native getter states that the local Data Model fixture cannot reproduce.
[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableSnapshotFieldFailureTests
{
    [Fact]
    public void RegularField_PositionUnavailable_ReportsWhyFiltersAreMissing()
    {
        var field = FakeField.Create<Excel.PivotField>();

        var result = PivotTableCommands.ReadRegularFieldSelection(
            field, "Pivot1", "book.xlsx", CancellationToken.None, 10);

        Assert.NotNull(result);
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
        Assert.False(result.LayoutAvailable);
        Assert.Equal("Position unavailable.", result.LayoutError);
        Assert.False(result.FiltersAvailable);
        Assert.Equal("The field identity or placement could not be read.", result.FiltersError);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(true)]
    [InlineData(false)]
    public void OlapCaptionFilter_MemberPropertyIdentity_DeterminesCriteriaAvailability(bool? memberProperty)
    {
        var filter = FakeField.Create<Excel.PivotFilter>(name => name switch
        {
            "get_FilterType" => Excel.XlPivotFilterType.xlCaptionEquals,
            "get_IsMemberPropertyFilter" => memberProperty ??
#pragma warning disable CA2201 // Simulate unavailable member-property metadata.
                throw new COMException("Member-property metadata unavailable."),
#pragma warning restore CA2201
            "get_Value1" => "Red",
            _ => throw new NotSupportedException(name)
        });
        var cube = CreateOlapHierarchy(filter);

        var result = PivotTableCommands.ReadOlapFieldSelection(
            cube, "Pivot1", "book.xlsx", CancellationToken.None, 10, null);

        Assert.NotNull(result);
        Assert.True(result.SelectionStateAvailable);
        Assert.Equal(memberProperty == false, result.FiltersAvailable);
        if (memberProperty == false)
        {
            Assert.Null(result.FiltersError);
            var criterion = Assert.Single(result.Filters);
            Assert.Equal("[Product].[Product].[Product]", criterion.LevelUniqueName);
            Assert.Equal("CaptionEquals", criterion.Type);
            Assert.Equal("Red", criterion.Value1);
        }
        else
        {
            Assert.Empty(result.Filters);
            Assert.Contains("member-property", result.FiltersError, StringComparison.OrdinalIgnoreCase);
        }
    }

    public static TheoryData<string, int> FilterGetterFailures
    {
        get
        {
            var cases = new TheoryData<string, int>();
            foreach (var getter in new[] { "get_FilterType", "get_IsMemberPropertyFilter", "get_Value1", "get_Value2" })
            {
                foreach (int hresult in new[]
                {
                    ResiliencePipelines.RPC_E_CALL_FAILED,
                    ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE,
                    ResiliencePipelines.RPC_E_DISCONNECTED,
                    unchecked((int)0x800A03EC)
                })
                    cases.Add(getter, hresult);
            }
            return cases;
        }
    }

    [Theory]
    [MemberData(nameof(FilterGetterFailures))]
    public void OlapFilter_GetterFailure_PreservesFatalErrorsAndReportsUnavailableMetadata(string getter, int hresult)
    {
#pragma warning disable CA2201 // Inject native getter failures without disconnecting a real Excel session.
        var failure = new COMException("Filter getter failed.", hresult);
#pragma warning restore CA2201
        var filter = FakeField.Create<Excel.PivotFilter>(name => name == getter ? throw failure : name switch
        {
            "get_FilterType" => Excel.XlPivotFilterType.xlDateBetween,
            "get_IsMemberPropertyFilter" => false,
            "get_Value1" => new DateTime(2024, 2, 10).ToOADate(),
            "get_Value2" => new DateTime(2024, 3, 5).ToOADate(),
            _ => throw new NotSupportedException(name)
        });
        var cube = CreateOlapHierarchy(filter);

        if (hresult != unchecked((int)0x800A03EC))
        {
            var thrown = Assert.Throws<COMException>(() => PivotTableCommands.ReadOlapFieldSelection(
                cube, "Pivot1", "book.xlsx", CancellationToken.None, 10, null));
            Assert.Same(failure, thrown);
        }
        else
        {
            var result = PivotTableCommands.ReadOlapFieldSelection(
                cube, "Pivot1", "book.xlsx", CancellationToken.None, 10, null);
            Assert.NotNull(result);
            Assert.True(result.SelectionStateAvailable);
            Assert.False(result.FiltersAvailable);
            Assert.Empty(result.Filters);
            Assert.False(string.IsNullOrWhiteSpace(result.FiltersError));
            Assert.Contains(failure.Message, result.FiltersError, StringComparison.Ordinal);
            Assert.Contains($"0x{failure.HResult:X8}", result.FiltersError, StringComparison.Ordinal);
        }
    }

    private static Excel.CubeField CreateOlapHierarchy(Excel.PivotFilter filter)
    {
        var filters = FakeField.Create<Excel.PivotFilters>(name => name switch
        {
            "get_Count" => 1,
            "get__Default" => filter,
            _ => throw new NotSupportedException(name)
        });
        var field = FakeField.Create<Excel.PivotField>(name => name switch
        {
            "get_Name" => "[Product].[Product].[Product]",
            "get_Orientation" => Excel.XlPivotFieldOrientation.xlRowField,
            "get_PivotFilters" => filters,
            _ => throw new NotSupportedException(name)
        });
        var levels = FakeField.Create<Excel.PivotFields>(name => name switch
        {
            "get_Count" => 1,
            "Item" => field,
            _ => throw new NotSupportedException(name)
        });
        return FakeField.Create<Excel.CubeField>(name => name switch
        {
            "get_Name" or "get_Caption" => "[Product].[Product]",
            "get_Orientation" => Excel.XlPivotFieldOrientation.xlRowField,
            "get_Position" => 1,
            "get_AllItemsVisible" => true,
            "get_PivotFields" => levels,
            _ => throw new NotSupportedException(name)
        });
    }

    [Fact]
    public void OlapHierarchy_PositionUnavailable_ReportsWhyFiltersAreMissing()
    {
        var cubeField = FakeField.Create<Excel.CubeField>();

        var result = PivotTableCommands.ReadOlapFieldSelection(
            cubeField, "Pivot1", "book.xlsx", CancellationToken.None, 10, null);

        Assert.NotNull(result);
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
        Assert.False(result.LayoutAvailable);
        Assert.Equal("Position unavailable.", result.LayoutError);
        Assert.False(result.FiltersAvailable);
        Assert.Equal("The hierarchy identity or placement could not be read.", result.FiltersError);
    }

    public class FakeField : DispatchProxy
    {
        // Core reads SourceName late-bound, which binds to this public property.
        public string SourceName { get; set; } = "Region";

        public Func<string, object?>? ReadMember { get; set; }

        public static T Create<T>(Func<string, object?>? readMember = null) where T : class
        {
            // Embedded PIAs expose different member subsets; implement Core's actual interface.
            var interfaceType = typeof(PivotTableCommands).Assembly.GetType(typeof(T).FullName!) ?? typeof(T);
            var proxy = DispatchProxy.Create(interfaceType, typeof(FakeField));
            ((FakeField)proxy).ReadMember = readMember;
            return (T)proxy;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            ReadMember is not null ? ReadMember(targetMethod!.Name) : targetMethod!.Name switch
            {
                "get_SourceName" or "get_Name" or "get_Caption" => "Region",
                "get_Orientation" => Excel.XlPivotFieldOrientation.xlRowField,
#pragma warning disable CA2201 // Simulate an Excel COM getter failure without starting Excel.
                "get_Position" => throw new COMException("Position unavailable."),
#pragma warning restore CA2201
                _ => throw new NotSupportedException(targetMethod.Name)
            };
    }
}
