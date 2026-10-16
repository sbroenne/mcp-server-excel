using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

// Inject getter failures that cannot be reproduced deterministically in a real workbook.
[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ErrorHandling")]
[Trait("RequiresExcel", "false")]
public sealed class NativePropertyFailureTests
{
    [Theory]
    [InlineData(typeof(ConnectionCommands), "GetBackgroundQuerySetting")]
    [InlineData(typeof(ConnectionCommands), "GetRefreshOnFileOpenSetting")]
    [InlineData(typeof(ConnectionCommands), "GetSavePasswordSetting")]
    [InlineData(typeof(ConnectionCommands), "GetRefreshPeriod")]
    [InlineData(typeof(ConnectionCommands), "GetConnectionString")]
    [InlineData(typeof(ConnectionCommands), "GetCommandText")]
    [InlineData(typeof(ConnectionCommands), "GetCommandType")]
    [InlineData(typeof(ConditionalFormattingCommands), "ReadRuleTypeInt")]
    [InlineData(typeof(DataModelCommands), "TryGetAdoConnectionString")]
    [InlineData(typeof(Sbroenne.ExcelMcp.Core.Commands.Screenshot.ScreenshotCommands), "GetActiveSheet")]
    [InlineData(typeof(Sbroenne.ExcelMcp.Core.Commands.Screenshot.ScreenshotCommands), "BringExcelToForeground")]
    [InlineData(typeof(Sbroenne.ExcelMcp.Core.Commands.Screenshot.ScreenshotCommands), "MeasurePaneOrigin")]
    public void UnexpectedComGetterFailurePropagatesWithItsOriginalHResult(Type owner, string methodName)
    {
#pragma warning disable CA2201 // Inject Excel's ambiguous error rather than treating it as unsupported.
        var failure = new COMException("Native getter failed.", unchecked((int)0x800A03EC));
#pragma warning restore CA2201
        object target = owner == typeof(ConditionalFormattingCommands)
            ? new FailingRule(failure)
            : new FailingProperties(failure);
        var method = owner.GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic)!;

        var wrapper = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, [target]));
        Assert.Same(failure, wrapper.InnerException);
    }

    public sealed class FailingProperties(COMException failure)
    {
        public int Type { get; } = 1;
        public object OLEDBConnection => throw failure;
        public object ConnectionString => throw failure;
        public object ActiveSheet => throw failure;
        public long Hwnd => throw failure;
        public int ScrollRow { set => throw failure; }
    }

    public sealed class FailingRule(COMException failure)
    {
        public int Type => throw failure;
    }
}
