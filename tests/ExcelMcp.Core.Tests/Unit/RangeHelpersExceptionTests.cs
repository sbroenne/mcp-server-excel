using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class RangeHelpersExceptionTests
{
    [Fact]
    public void ResolveRange_WhenNamedRangeLookupThrowsUnexpectedError_PropagatesError()
    {
        var expected = new InvalidOperationException("Unexpected lookup failure.");

        var actual = Assert.Throws<InvalidOperationException>(() =>
            RangeHelpers.ResolveRange(
                ExcelInterfaceProxy.Create<Excel.Workbook>((_, _) => throw expected),
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
                CreateWorkbook(comError),
                string.Empty,
                "MissingRange",
                out _));

        Assert.Equal(OperationFailureCategory.NotFound, actual.ErrorCategory);
        Assert.Same(comError, actual.InnerException);
    }

    [Fact]
    public void ResolveRange_WhenOnlyWorksheetScopedNameMatches_ThrowsCategorizedNotFound()
    {
#pragma warning disable CA2201 // Synthetic COM exception exercises pure classification behavior.
        var comError = new COMException("Unknown name", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        var actual = Assert.Throws<OperationFailureException>(() =>
            RangeHelpers.ResolveRange(
                CreateWorkbook(comError, resolvedName: "Sheet1!MissingRange"),
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
                CreateWorkbook(null, resolvedName: "FormulaName", rangeFailure: expected),
                string.Empty,
                "FormulaName",
                out _));

        Assert.Same(expected, actual);
    }

    [Fact]
    public void ResolveRange_WhenNameEnumerationFails_PropagatesError()
    {
#pragma warning disable CA2201 // Synthetic COM exception exercises pure classification behavior.
        var lookupException = new COMException("Unknown name", unchecked((int)0x800A03EC));
        var expected = new COMException("Name enumeration failed.", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        var actual = Assert.Throws<COMException>(() =>
            RangeHelpers.ResolveRange(
                CreateWorkbook(lookupException, enumerationFailure: expected),
                string.Empty,
                "MissingRange",
                out _));

        Assert.Same(expected, actual);
    }

    private static Excel.Workbook CreateWorkbook(
        Exception? lookupFailure, string? resolvedName = null,
        Exception? rangeFailure = null, Exception? enumerationFailure = null)
    {
        var name = ExcelInterfaceProxy.Create<Excel.Name>((member, _) => member.Name switch
        {
            "get_Name" => resolvedName,
            "get_RefersToRange" => throw rangeFailure!,
            _ => throw new InvalidOperationException($"Unexpected name read: {member.Name}")
        });
        var names = ExcelInterfaceProxy.Create<Excel.Names>((member, args) =>
        {
            if (member.Name == "get_Count")
                return resolvedName is not null || enumerationFailure is not null ? 1 : 0;
            if (member.Name == "Item")
            {
                if (args![0] is string && lookupFailure is not null)
                    throw lookupFailure;
                if (args[0] is int && enumerationFailure is not null)
                    throw enumerationFailure;
                return name;
            }
            throw new InvalidOperationException($"Unexpected names read: {member.Name}");
        });
        return ExcelInterfaceProxy.Create<Excel.Workbook>((member, _) => member.Name == "get_Names"
            ? names : throw new InvalidOperationException($"Unexpected workbook read: {member.Name}"));
    }
}
