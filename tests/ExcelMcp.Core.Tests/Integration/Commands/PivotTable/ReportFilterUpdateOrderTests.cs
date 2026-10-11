using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PivotTable;

// Service cannot deterministically inject provider normalization during ManualUpdate restoration.
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Medium")]
public sealed class ReportFilterUpdateOrderTests(TempDirectoryFixture fixture) :
    IClassFixture<TempDirectoryFixture>
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [InlineData(true, true)]
    public void SetReportFilter_ReadsSelectionAfterRestoringUpdates(
        bool restorationFails, bool mutationFails = false)
    {
        using var inner = ExcelSession.BeginBatch(fixture.CreateTestFile());
        using var batch = DispatchProxy.Create<IExcelBatch, ProviderProxy>();
        var provider = (ProviderProxy)batch;
        provider.Inner = inner;
        provider.RestorationFails = restorationFails;
        provider.MutationFails = mutationFails;

        var result = new PivotTableCommands().SetReportFilter(
            batch, "Sheet", "Pivot", "[Region]", ["[Requested]"]);

        Assert.True(provider.RestorationAttempted, result.ErrorMessage);
        Assert.False(result.ShowAll);
        Assert.Equal(["[Normalized]"], result.SelectedItems);
        Assert.Equal(!restorationFails, result.Success);
        Assert.Equal(restorationFails, result.MayHavePartiallyChanged);
        Assert.False(result.RollbackAttempted);
        if (restorationFails)
            Assert.Contains("Injected restoration failure", result.ErrorMessage);
        else
            Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        if (mutationFails)
            Assert.Contains("Injected mutation failure", result.ErrorMessage);
    }

    public class ProviderProxy : DispatchProxy
    {
        public ProviderProxy Owner { get; set; } = null!;
        public IExcelBatch Inner { get; set; } = null!;
        public Type ExcelType { get; set; } = typeof(IExcelBatch);
        public bool RestorationFails { get; set; }
        public bool MutationFails { get; set; }
        public bool RestorationAttempted { get; set; }
        public bool ManualUpdate { get; set; }
        public string[] Selection { get; set; } = [];
        public Excel.CubeFields CubeFields => Create<Excel.CubeFields>(Owner);
        public int Count { get; } = 1;

        private static T Create<T>(ProviderProxy owner) where T : class
        {
            // Embedded PIAs expose different member subsets; implement Core's actual interface.
            var interfaceType = typeof(PivotTableCommands).Assembly.GetType(typeof(T).FullName!) ?? typeof(T);
            var value = (T)DispatchProxy.Create(interfaceType, typeof(ProviderProxy));
            var proxy = (ProviderProxy)(object)value;
            proxy.Owner = owner;
            proxy.ExcelType = typeof(T);
            return value;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            var name = targetMethod!.Name;
            if (ExcelType == typeof(IExcelBatch))
            {
                if (name == nameof(IDisposable.Dispose)) return null;
                if (name == "get_WorkbookPath") return "synthetic.xlsx";
                if (name == nameof(IExcelBatch.Execute))
                {
                    Owner = this;
                    var workbook = Create<Excel.Workbook>(this);
                    return Inner.Execute((context, token) =>
                    {
                        try
                        {
                            return ((Delegate)args![0]!).DynamicInvoke(
                                new ExcelContext(context.WorkbookPath, context.App, workbook), token);
                        }
                        catch (TargetInvocationException ex) when (ex.InnerException is not null)
                        {
                            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                            throw;
                        }
                    });
                }
            }
            if (ExcelType == typeof(Excel.Workbook) && name == "get_Worksheets")
                return Create<Excel.Sheets>(Owner);
            if (ExcelType == typeof(Excel.Sheets) && name == "get__Default")
                return Create<Excel.Worksheet>(Owner);
            if (ExcelType == typeof(Excel.Worksheet) && name == "PivotTables")
                return Create<Excel.PivotTables>(Owner);
            if (ExcelType == typeof(Excel.PivotTables) && name == "Item")
                return Create<Excel.PivotTable>(Owner);
            if (ExcelType == typeof(Excel.PivotTable))
            {
                if (name == "get_CubeFields") return CubeFields;
                if (name == "get_ManualUpdate") return Owner.ManualUpdate;
                if (name == "set_ManualUpdate")
                {
                    Owner.ManualUpdate = (bool)args![0]!;
                    if (!Owner.ManualUpdate)
                    {
                        Owner.RestorationAttempted = true;
                        Owner.Selection = ["[Normalized]"];
                        if (Owner.RestorationFails)
                        {
#pragma warning disable CA2201 // Simulate a provider failure after applying its normalized selection.
                            throw new COMException("Injected restoration failure", unchecked((int)0x800A03EC));
#pragma warning restore CA2201
                        }
                    }
                    return null;
                }
            }
            if (ExcelType == typeof(Excel.CubeFields))
            {
                if (name == "get_Count") return 1;
                if (name == "get__Default") return Create<Excel.CubeField>(Owner);
            }
            if (ExcelType == typeof(Excel.CubeField))
            {
                if (name == "get_Orientation") return Excel.XlPivotFieldOrientation.xlPageField;
                if (name == "get_PivotFields") return Create<Excel.PivotFields>(Owner);
                if (name == "set_EnableMultiplePageItems") return null;
                if (name == "get_EnableMultiplePageItems") return true;
                if (name == "get_AllItemsVisible") return false;
            }
            if (ExcelType == typeof(Excel.PivotFields))
            {
                if (name == "get_Count") return 1;
                if (name == "Item") return Create<Excel.PivotField>(Owner);
            }
            if (ExcelType == typeof(Excel.PivotField))
            {
                if (name == "get_VisibleItemsList") return Owner.Selection;
                if (name == "set_VisibleItemsList")
                {
                    Owner.Selection = (string[])args![0]!;
                    if (Owner.MutationFails)
                        throw new InvalidOperationException("Injected mutation failure");
                    return null;
                }
            }
            throw new NotSupportedException($"Unexpected {ExcelType.Name} member: {name}");
        }
    }
}
