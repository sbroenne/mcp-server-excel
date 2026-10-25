using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "false")]
public sealed class DataModelRecordsetCleanupTests
{
    [Theory]
    [InlineData((int)ADODB.ObjectStateEnum.adStateOpen)]
    [InlineData((int)(ADODB.ObjectStateEnum.adStateOpen | ADODB.ObjectStateEnum.adStateFetching))]
    public void OpenRecordsetIsClosed(int state)
    {
        var recordset = new Recordset { State = state };

        InvokeClose(recordset, null);

        Assert.True(recordset.Closed);
    }

    [Fact]
    public void ClosedRecordsetIsNotClosedAgain()
    {
        var recordset = new Recordset { State = (int)ADODB.ObjectStateEnum.adStateClosed };

        InvokeClose(recordset, null);

        Assert.False(recordset.CloseAttempted);
    }

    [Fact]
    public void CleanupFailureIsNotReportedAsSuccess()
    {
        var recordset = new Recordset { CloseFailure = CreateFailure() };

        var wrapper = Assert.Throws<TargetInvocationException>(() => InvokeClose(recordset, null));

        Assert.Same(recordset.CloseFailure, wrapper.InnerException);
        Assert.True(recordset.CloseAttempted);
    }

    [Fact]
    public void PrimaryAndCleanupFailuresAreBothPreserved()
    {
        var primary = new InvalidOperationException("Reading records failed.");
        var recordset = new Recordset { CloseFailure = CreateFailure() };

        var wrapper = Assert.Throws<TargetInvocationException>(() => InvokeClose(recordset, primary));

        var aggregate = Assert.IsType<AggregateException>(wrapper.InnerException);
        Assert.Equal([primary, recordset.CloseFailure], aggregate.InnerExceptions);
        Assert.True(recordset.CloseAttempted);
    }

    [Fact]
    public void RecordsetCleanupPreservesCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var primary = new OperationCanceledException(cancellation.Token);
        var recordset = new Recordset { CloseFailure = CreateFailure() };

        var wrapper = Assert.Throws<TargetInvocationException>(() =>
            InvokeClose(recordset, primary));

        var canceled = Assert.IsType<OperationCanceledException>(wrapper.InnerException);
        Assert.Equal(cancellation.Token, canceled.CancellationToken);
        var failures = Assert.IsType<AggregateException>(canceled.InnerException);
        Assert.Equal([primary, recordset.CloseFailure], failures.InnerExceptions);
    }

    private static void InvokeClose(Recordset recordset, Exception? primary)
    {
        var method = typeof(DataModelCommands).GetMethod("CloseRecordset",
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);
        var typed = ExcelInterfaceProxy.Create<ADODB._Recordset>((member, _) =>
        {
            if (member.Name == "get_State")
                return recordset.State;
            if (member.Name == "Close")
            {
                recordset.Close();
                return null;
            }
            throw new NotSupportedException(member.Name);
        });
        method.Invoke(null, [typed, primary]);
    }

#pragma warning disable CA2201 // Simulate the exact native close failure without starting Excel.
    private static COMException CreateFailure() =>
        new("Closing the native recordset failed.", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

    public sealed class Recordset
    {
        public int State { get; init; } = (int)ADODB.ObjectStateEnum.adStateOpen;
        public COMException? CloseFailure { get; init; }
        public bool Closed { get; private set; }
        public bool CloseAttempted { get; private set; }

        public void Close()
        {
            CloseAttempted = true;
            if (CloseFailure is not null)
                throw CloseFailure;
            Closed = true;
        }
    }
}
