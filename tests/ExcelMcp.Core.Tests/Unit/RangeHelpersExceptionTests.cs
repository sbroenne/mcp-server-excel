using System.Dynamic;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("Speed", "Fast")]
public class RangeHelpersExceptionTests
{
    [Fact]
    public void ResolveRange_WhenNamedRangeLookupThrowsUnexpectedError_PropagatesError()
    {
        var expected = new InvalidOperationException("Unexpected lookup failure.");

        var actual = Assert.Throws<InvalidOperationException>(() =>
            RangeHelpers.ResolveRange(
                new ThrowingWorkbook(expected),
                string.Empty,
                "TestRange",
                out _));

        Assert.Same(expected, actual);
    }

    [Fact]
    public void ResolveRange_WhenNamedRangeDoesNotExist_ThrowsCategorizedNotFound()
    {
#pragma warning disable CA2201 // Synthetic COM exception exercises pure classification behavior.
        var comError = new COMException("Unknown name", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        var actual = Assert.Throws<OperationFailureException>(() =>
            RangeHelpers.ResolveRange(
                new ThrowingWorkbook(comError),
                string.Empty,
                "MissingRange",
                out _));

        Assert.Equal(OperationFailureCategory.NotFound, actual.ErrorCategory);
        Assert.Same(comError, actual.InnerException);
    }

    private sealed class ThrowingWorkbook(Exception exception) : DynamicObject
    {
        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            throw exception;
        }
    }
}
