using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PivotTable;

// Service cannot inject a cancelled token into an already-running STA callback.
[Collection("Sequential")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Medium")]
public sealed class ReportFilterCancellationTests(TempDirectoryFixture fixture) :
    IClassFixture<TempDirectoryFixture>
{
    [Fact]
    public void SetReportFilter_CancelledCallback_PropagatesBeforeWorkbookAccess()
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.CreateTestFile());
        var savedBefore = innerBatch.Execute((context, _) => context.Book.Saved);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var batch = new InjectedCancellationBatch(innerBatch, cancellation.Token);
        var commands = new PivotTableCommands();

        var error = Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.SetReportFilter(batch, "MissingSheet", "MissingPivot",
                "[Sales].[Region]", []));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.False(innerBatch.HasTimedOutOperation);
        Assert.Equal(savedBefore, innerBatch.Execute((context, _) => context.Book.Saved));
    }

    [Theory]
    [InlineData(ResiliencePipelines.RPC_E_CALL_FAILED)]
    [InlineData(ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE)]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    [InlineData(ResiliencePipelines.RPC_E_CALL_FAILED, true)]
    [InlineData(ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE, true)]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED, true)]
    public void SetReportFilter_DisconnectedWorkbook_PropagatesOriginalFailure(
        int hresult, bool disconnectAtCubeFields = false)
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.CreateTestFile());
        using var batch = DispatchProxy.Create<IExcelBatch, DisconnectedBatch>();
        var proxy = (DisconnectedBatch)batch;
        proxy.Inner = innerBatch;
        proxy.DisconnectAtCubeFields = disconnectAtCubeFields;
        // Excel disconnects nondeterministically; inject its COM error at the selected read.
#pragma warning disable CA2201
        proxy.Failure = new COMException("Injected Excel disconnect.", hresult);
#pragma warning restore CA2201

        var error = Assert.Throws<COMException>(() =>
            new PivotTableCommands().SetReportFilter(batch, "Sheet", "Pivot", "[Sales].[Region]", []));

        Assert.Same(proxy.Failure, error);
        Assert.True(proxy.WorkbookReadAttempted);
        Assert.Equal(disconnectAtCubeFields, proxy.CubeFieldsReadAttempted);
    }

    public class DisconnectedWorkbook : DispatchProxy
    {
        public DisconnectedBatch Owner { get; set; } = null!;

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name == "get_Worksheets")
            {
                Owner.WorkbookReadAttempted = true;
                if (Owner.DisconnectAtCubeFields)
                    return DisconnectedPivotObjects.Create<Excel.Sheets>(Owner);
                throw Owner.Failure;
            }
            throw new NotSupportedException($"Unexpected workbook member: {targetMethod?.Name}");
        }
    }

    public class DisconnectedPivotObjects : DispatchProxy
    {
        public DisconnectedBatch Owner { get; set; } = null!;
        public Type ExcelType { get; set; } = null!;

        public Excel.CubeFields CubeFields
        {
            get
            {
                Owner.CubeFieldsReadAttempted = true;
                throw Owner.Failure;
            }
        }

        public static T Create<T>(DisconnectedBatch owner) where T : class
        {
            // Match Core's embedded PIA member subset, including PivotCache().
            var interfaceType = typeof(PivotTableCommands).Assembly.GetType(typeof(T).FullName!) ?? typeof(T);
            var value = (T)DispatchProxy.Create(interfaceType, typeof(DisconnectedPivotObjects));
            var proxy = (DisconnectedPivotObjects)(object)value;
            proxy.Owner = owner;
            proxy.ExcelType = typeof(T);
            return value;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (ExcelType == typeof(Excel.Sheets) && targetMethod?.Name == "get__Default")
                return Create<Excel.Worksheet>(Owner);
            if (ExcelType == typeof(Excel.Worksheet) && targetMethod?.Name == "PivotTables")
                return Create<Excel.PivotTables>(Owner);
            if (ExcelType == typeof(Excel.PivotTables) && targetMethod?.Name == "Item")
                return Create<Excel.PivotTable>(Owner);
            if (ExcelType == typeof(Excel.PivotTable) && targetMethod?.Name == "PivotCache")
            {
                Owner.CacheReadCount++;
                if (Owner.FailCacheRead)
                    throw Owner.Failure;
                return Create<Excel.PivotCache>(Owner);
            }
            if (ExcelType == typeof(Excel.PivotCache) && targetMethod?.Name == "get_OLAP")
                return Owner.Olap;
            if (ExcelType == typeof(Excel.PivotTable) && targetMethod?.Name == "get_CubeFields")
                return CubeFields;
            throw new NotSupportedException($"Unexpected {ExcelType.Name} member: {targetMethod?.Name}");
        }
    }

    public class DisconnectedBatch : DispatchProxy
    {
        public IExcelBatch Inner { get; set; } = null!;
        public COMException Failure { get; set; } = null!;
        public bool WorkbookReadAttempted { get; set; }
        public bool DisconnectAtCubeFields { get; set; }
        public bool CubeFieldsReadAttempted { get; set; }
        public bool Olap { get; set; } = true;
        public bool FailCacheRead { get; set; }
        public int CacheReadCount { get; set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name == nameof(IDisposable.Dispose))
                return null;
            if (targetMethod?.Name == "get_WorkbookPath")
                return Inner.WorkbookPath;
            if (targetMethod?.Name != nameof(IExcelBatch.Execute))
                throw new NotSupportedException($"Unexpected batch member: {targetMethod?.Name}");

            Exception? failure = null;
            object? result = null;
            Inner.Execute((context, token) =>
            {
                var workbook = DispatchProxy.Create<Excel.Workbook, DisconnectedWorkbook>();
                ((DisconnectedWorkbook)workbook).Owner = this;
                try
                {
                    var callback = (Delegate)args![0]!;
                    result = callback.DynamicInvoke(
                        new ExcelContext(context.WorkbookPath, context.App, workbook), token);
                }
                catch (TargetInvocationException ex) when (ex.InnerException is not null)
                {
                    failure = ex.InnerException;
                }
            });
            if (failure is not null)
                ExceptionDispatchInfo.Capture(failure).Throw();
            return result;
        }
    }
}
