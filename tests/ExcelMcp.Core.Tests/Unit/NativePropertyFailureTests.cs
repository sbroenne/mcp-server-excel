using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
        object target = methodName switch
        {
            "ReadRuleTypeInt" => ExcelInterfaceProxy.Create<Excel.FormatCondition>((_, _) => throw failure),
            "TryGetAdoConnectionString" => ExcelInterfaceProxy.Create<ADODB._Connection>((_, _) => throw failure),
            "GetActiveSheet" or "BringExcelToForeground" =>
                ExcelInterfaceProxy.Create<Excel.Application>((_, _) => throw failure),
            "MeasurePaneOrigin" => ExcelInterfaceProxy.Create<Excel.Window>((_, _) => throw failure),
            _ => ExcelInterfaceProxy.Create<Excel.WorkbookConnection>((member, _) =>
                member.Name == "get_Type" ? Excel.XlConnectionType.xlConnectionTypeOLEDB : throw failure)
        };
        var method = owner.GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic)!;

        var wrapper = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, [target]));
        Assert.Same(failure, wrapper.InnerException);
    }

}
