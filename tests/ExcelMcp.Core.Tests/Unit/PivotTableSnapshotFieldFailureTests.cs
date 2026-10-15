using System.Reflection;
using System.Runtime.InteropServices;
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
        var cube = FakeField.Create<Excel.CubeField>(name => name switch
        {
            "get_Name" or "get_Caption" => "[Product].[Product]",
            "get_Orientation" => Excel.XlPivotFieldOrientation.xlRowField,
            "get_Position" => 1,
            "get_AllItemsVisible" => true,
            "get_PivotFields" => levels,
            _ => throw new NotSupportedException(name)
        });

        var result = PivotTableCommands.ReadOlapFieldSelection(
            cube, "Pivot1", "book.xlsx", CancellationToken.None, 10);

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

    [Fact]
    public void OlapHierarchy_PositionUnavailable_ReportsWhyFiltersAreMissing()
    {
        var cubeField = FakeField.Create<Excel.CubeField>();

        var result = PivotTableCommands.ReadOlapFieldSelection(
            cubeField, "Pivot1", "book.xlsx", CancellationToken.None, 10);

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
