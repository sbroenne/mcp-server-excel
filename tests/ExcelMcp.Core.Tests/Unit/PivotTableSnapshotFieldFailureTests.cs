using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

// Excel cannot be made to fail a field Position read on demand, so a fake field injects it.
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

        public static T Create<T>() where T : class
        {
            // Embedded PIAs expose different member subsets; implement Core's actual interface.
            var interfaceType = typeof(PivotTableCommands).Assembly.GetType(typeof(T).FullName!) ?? typeof(T);
            return (T)DispatchProxy.Create(interfaceType, typeof(FakeField));
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            targetMethod!.Name switch
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
