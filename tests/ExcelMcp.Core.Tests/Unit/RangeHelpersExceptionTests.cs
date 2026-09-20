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
                new Workbook(new MissingNames(comError)),
                string.Empty,
                "MissingRange",
                out _));

        Assert.Equal(OperationFailureCategory.NotFound, actual.ErrorCategory);
        Assert.Same(comError, actual.InnerException);
    }

    [Fact]
    public void ResolveRange_WhenNamedRangeCannotResolveToRange_PropagatesError()
    {
#pragma warning disable CA2201 // Synthetic COM exception exercises pure classification behavior.
        var expected = new COMException("Name does not refer to a range.", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        var actual = Assert.Throws<COMException>(() =>
            RangeHelpers.ResolveRange(
                new Workbook(new NamesWithFormulaName(expected)),
                string.Empty,
                "FormulaName",
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

    private sealed class Workbook(object names) : DynamicObject
    {
        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            result = binder.Name == "Names" ? names : null;
            return binder.Name == "Names";
        }
    }

    private sealed class MissingNames(COMException exception) : DynamicObject
    {
        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            result = binder.Name == "Count" ? 0 : null;
            return binder.Name == "Count";
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
        {
            if (binder.Name == "Item")
            {
                throw exception;
            }

            result = null;
            return false;
        }
    }

    private sealed class NamesWithFormulaName(COMException exception) : DynamicObject
    {
        private readonly FormulaName _name = new(exception);

        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            result = binder.Name == "Count" ? 1 : null;
            return binder.Name == "Count";
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
        {
            if (binder.Name == "Item" && args is [string])
            {
                result = _name;
                return true;
            }

            if (binder.Name == "Item" && args is [int])
            {
                result = _name;
                return true;
            }

            result = null;
            return false;
        }
    }

    private sealed class FormulaName(COMException exception) : DynamicObject
    {
        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            if (binder.Name == "Name")
            {
                result = "FormulaName";
                return true;
            }

            if (binder.Name == "RefersToRange")
            {
                throw exception;
            }

            result = null;
            return false;
        }
    }
}
