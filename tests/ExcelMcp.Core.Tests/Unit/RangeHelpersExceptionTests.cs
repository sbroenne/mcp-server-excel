using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
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
                ExcelProxy.Create<Excel.Workbook>((_, _) => throw expected),
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
                CreateWorkbook(0, _ => throw comError),
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
                CreateWorkbook(1, key => key is string
                    ? throw comError
                    : ExcelProxy.Create<Excel.Name>((member, _) => member == "get_Name"
                        ? "Sheet1!MissingRange" : throw new NotSupportedException(member))),
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
                CreateWorkbook(1, _ => ExcelProxy.Create<Excel.Name>((member, _) => member switch
                {
                    "get_Name" => "FormulaName",
                    "get_RefersToRange" => throw expected,
                    _ => throw new NotSupportedException(member)
                })),
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
                CreateWorkbook(1, key => key is string ? throw lookupException : throw expected),
                string.Empty,
                "MissingRange",
                out _));

        Assert.Same(expected, actual);
    }

    private static Excel.Workbook CreateWorkbook(int count, Func<object?, Excel.Name> item)
    {
        var names = ExcelProxy.Create<Excel.Names>((member, args) => member switch
        {
            "get_Count" => count,
            "Item" => item(args![0]),
            _ => throw new NotSupportedException(member)
        });
        return ExcelProxy.Create<Excel.Workbook>((member, _) => member == "get_Names"
            ? names : throw new NotSupportedException(member));
    }

    public class ExcelProxy : DispatchProxy
    {
        private Func<string, object?[]?, object?> _invoke = null!;

        public static T Create<T>(Func<string, object?[]?, object?> invoke) where T : class
        {
            var interfaceType = typeof(RangeHelpers).Assembly.GetType(typeof(T).FullName!) ?? typeof(T);
            var proxy = DispatchProxy.Create(interfaceType, typeof(ExcelProxy));
            ((ExcelProxy)proxy)._invoke = invoke;
            return (T)proxy;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            _invoke(targetMethod!.Name, args);
    }
}
