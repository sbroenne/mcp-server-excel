using System.Dynamic;
using Sbroenne.ExcelMcp.Core.Commands.Range;
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

    private sealed class ThrowingWorkbook(Exception exception) : DynamicObject
    {
        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            throw exception;
        }
    }
}
