using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "false")]
public sealed class DataModelRecordsetCleanupTests
{
    [Fact]
    public void OpenRecordsetIsClosedWhenStateMarshalsAsDouble()
    {
        var recordset = new Recordset();

        InvokeClose(recordset, null);

        Assert.True(recordset.Closed);
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

    private static void InvokeClose(Recordset recordset, Exception? primary)
    {
        var method = typeof(DataModelCommands).GetMethod("CloseRecordset",
            BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);
        method.Invoke(null, [recordset, primary]);
    }

#pragma warning disable CA2201 // Simulate the exact native close failure without starting Excel.
    private static COMException CreateFailure() =>
        new("Closing the native recordset failed.", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

    public sealed class Recordset
    {
        public double State { get; } = 1d;
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
