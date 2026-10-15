using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeThemeFailureTests
{
    [Fact]
    public void AmbiguousExcelFailureIsNotClassifiedAsANonThemeColor()
    {
        var failure = NativeFailure(unchecked((int)0x800A03EC));
        var wrapper = Assert.Throws<TargetInvocationException>(() => Read(() => throw failure));
        Assert.Same(failure, wrapper.InnerException);
    }

    [Theory]
    [InlineData(unchecked((int)0x80020003))]
    [InlineData(unchecked((int)0x80020006))]
    [InlineData(unchecked((int)0x80004001))]
    public void KnownUnavailableThemePropertyCanBeOmitted(int hresult)
    {
        Assert.Null(Read(() => throw NativeFailure(hresult)));
    }

    [Fact]
    public void AvailableThemeIndexIsPreserved()
    {
        Assert.Equal(5, Read(() => 5));
    }

    private static object? Read(Func<object> getter)
    {
        var method = typeof(RangeCommands).GetMethod("ReadFormatTheme", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        return method.Invoke(null, [getter, "font.themeColor", new List<string>()]);
    }

#pragma warning disable CA2201 // Simulate exact native property errors without starting Excel.
    private static COMException NativeFailure(int hresult) => new("Native color read failed.", hresult);
#pragma warning restore CA2201
}
