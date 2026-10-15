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
public sealed class ValidationAbsenceClassificationTests
{
    [Fact]
    public void ValidTypeDoesNotProbeForAbsence()
    {
        Assert.Equal(3, Read(() => 3, () => throw new InvalidOperationException("Unexpected probe.")));
    }

    [Fact]
    public void AmbiguousFailureWithExistingRulePreservesOriginalError()
    {
        var original = Failure(unchecked((int)0x800A03EC));
        var wrapper = Assert.Throws<TargetInvocationException>(() => Read(() => throw original, () => true));
        Assert.Same(original, wrapper.InnerException);
    }

    [Fact]
    public void AmbiguousFailureWithCorroboratedAbsenceReturnsNoType()
    {
        Assert.Null(Read(() => throw Failure(unchecked((int)0x800A03EC)), () => false));
    }

    [Theory]
    [InlineData(unchecked((int)0x80004005))]
    [InlineData(unchecked((int)0x80020003))]
    [InlineData(unchecked((int)0x80010108))]
    public void OtherNativeFailuresDoNotTriggerAbsenceProbe(int hresult)
    {
        var original = Failure(hresult);
        var wrapper = Assert.Throws<TargetInvocationException>(() =>
            Read(() => throw original, () => throw new InvalidOperationException("Unexpected probe.")));
        Assert.Same(original, wrapper.InnerException);
    }

    [Fact]
    public void CorroborationFailureIsNotClassifiedAsAbsence()
    {
        var probeError = Failure(unchecked((int)0x80004005));
        var wrapper = Assert.Throws<TargetInvocationException>(() =>
            Read(() => throw Failure(unchecked((int)0x800A03EC)), () => throw probeError));
        Assert.Same(probeError, wrapper.InnerException);
    }

    private static object? Read(Func<int> readType, Func<bool> containsValidation)
    {
        var method = typeof(RangeCommands).GetMethod("ReadValidationType", BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(method);
        return method.Invoke(null, [readType, containsValidation]);
    }

#pragma warning disable CA2201 // Simulate exact native validation failures without starting Excel.
    private static COMException Failure(int hresult) => new("Native validation read failed.", hresult);
#pragma warning restore CA2201
}
