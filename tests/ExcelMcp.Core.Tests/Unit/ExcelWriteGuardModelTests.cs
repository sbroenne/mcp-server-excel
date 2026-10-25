using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelWriteGuardModelTests
{
    [Theory]
    [InlineData((int)Excel.XlConnectionType.xlConnectionTypeMODEL, true, true, 1)]
    [InlineData((int)Excel.XlConnectionType.xlConnectionTypeOLEDB, false, true, 1)]
    [InlineData((int)Excel.XlConnectionType.xlConnectionTypeMODEL, true, false, 1)]
    [InlineData((int)Excel.XlConnectionType.xlConnectionTypeOLEDB, false, false, 1)]
    [InlineData((int)Excel.XlConnectionType.xlConnectionTypeMODEL, true, true, 2)]
    [InlineData((int)Excel.XlConnectionType.xlConnectionTypeOLEDB, false, true, 2)]
    public void ModelConnectionsDoNotChangeScreenUpdating(
        int connectionType, bool model, bool originalScreenUpdating, int bookCount)
    {
        bool screenUpdating = originalScreenUpdating;
        var writes = new List<bool>();
        var connection = Create<Excel.WorkbookConnection>((member, _) =>
            member.Name == "get_Type" ? Enum.ToObject(member.ReturnType, connectionType)
                : throw new NotSupportedException(member.Name));
        var connections = Create<Excel.Connections>((member, _) =>
            member.Name switch
            {
                "get_Count" => 1,
                "get_Item" or "get__Default" or "Item" => connection,
                _ => throw new NotSupportedException(member.Name)
            });
        var book = Create<Excel.Workbook>((member, _) =>
            member.Name == "get_Connections" ? connections
                : throw new NotSupportedException(member.Name));
        var emptyConnections = Create<Excel.Connections>((member, _) =>
            member.Name == "get_Count" ? 0 : throw new NotSupportedException(member.Name));
        var emptyBook = Create<Excel.Workbook>((member, _) =>
            member.Name == "get_Connections" ? emptyConnections
                : throw new NotSupportedException(member.Name));
        var books = Create<Excel.Workbooks>((member, args) =>
            member.Name switch
            {
                "get_Count" => bookCount,
                "get_Item" or "get__Default" or "Item" =>
                    Assert.IsType<int>(args![0]) == bookCount ? book : emptyBook,
                _ => throw new NotSupportedException(member.Name)
            });
        var app = Create<Excel.Application>((member, args) =>
        {
            if (member.Name == "get_Workbooks")
                return books;
            if (member.Name == "get_ScreenUpdating")
                return screenUpdating;
            if (member.Name == "set_ScreenUpdating")
            {
                screenUpdating = Assert.IsType<bool>(args![0]);
                writes.Add(screenUpdating);
                return null;
            }
            throw new NotSupportedException(member.Name);
        });

        using (var guard = new ExcelWriteGuard(app))
        {
            Assert.Equal(originalScreenUpdating && model, screenUpdating);
            using var nested = new ExcelWriteGuard(app);
            Assert.Equal(originalScreenUpdating && model, screenUpdating);
        }

        Assert.Equal(originalScreenUpdating, screenUpdating);
        Assert.Equal(model || !originalScreenUpdating ? Array.Empty<bool>() : [false, true], writes);

    }

    [Fact]
    public void FailedModelDetectionPreservesTheErrorAndDoesNotPoisonNesting()
    {
        var original = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800A03EC)));
        bool fail = true;
        bool screenUpdating = true;
        var books = Create<Excel.Workbooks>((member, _) =>
            member.Name == "get_Count" ? 0 : throw new NotSupportedException(member.Name));
        var app = Create<Excel.Application>((member, args) =>
        {
            if (member.Name == "get_Workbooks")
                return fail ? throw original : books;
            if (member.Name == "get_ScreenUpdating")
                return screenUpdating;
            if (member.Name == "set_ScreenUpdating")
            {
                screenUpdating = Assert.IsType<bool>(args![0]);
                return null;
            }
            throw new NotSupportedException(member.Name);
        });
        Assert.Same(original, Assert.Throws<COMException>(() => new ExcelWriteGuard(app)));
        Assert.True(screenUpdating);
        fail = false;
        using (var guard = new ExcelWriteGuard(app))
            Assert.False(screenUpdating);
        Assert.True(screenUpdating);
    }

    private static T Create<T>(Func<MethodInfo, object?[]?, object?> read) where T : class =>
        ExcelInterfaceProxy.Create<T>(read, typeof(ExcelWriteGuard).Assembly);
}
