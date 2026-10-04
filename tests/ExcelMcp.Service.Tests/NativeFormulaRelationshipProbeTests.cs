using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class NativeFormulaRelationshipProbeTests(
    PersistentServiceWorkbookFixture fixture, ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private static readonly string[] ProbeAddresses = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"];
    private static readonly string?[] DirectPrecedents = [null, "$A$1,$A$3", "$B$1", null, "$A$1", null, "$A$1", null];
    private static readonly string?[] DirectDependents = ["$B$1,$E$1,$G$1", "$C$1", null, null, null, null, null, null];
    private static readonly string?[] AllPrecedents = [null, "$A$1,$A$3", "$A$1:$B$1,$A$3", null, "$A$1", null, "$A$1", null];

    [Fact]
    public void NativeGetterBoundary_RecordsLocalRemoteDynamicAndLeafBehavior()
    {
        var sourceName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var otherName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sourceName, "A1:A3", [[2], [3], [4]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, otherName, "A1", [[7]]).Success);
        List<List<string>> formulas =
        [
            ["=SUM(A1,A3)", "=B1*2", $"='{otherName}'!A1", $"=A1+'{otherName}'!A1",
                "=INDIRECT(\"A2\")", "=OFFSET(A1,1,0)", "=1+2"]
        ];
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sourceName, "B1:H1", formulas).Success);
        _fixture.ExecuteRawVerification((context, ct) =>
        {
            Excel.Worksheet? originalSheet = null;
            Excel.Range? originalSelection = null;
            Excel.Worksheet? source = null;
            Excel.Worksheet? other = null;
            try
            {
                originalSheet = context.Book.ActiveSheet as Excel.Worksheet;
                originalSelection = context.App.Selection as Excel.Range;
                Assert.NotNull(originalSheet);
                source = ComUtilities.FindSheet(context.Book, sourceName);
                other = ComUtilities.FindSheet(context.Book, otherName);
                Assert.NotNull(source);
                Assert.NotNull(other);
                source.Activate();
                for (var index = 0; index < ProbeAddresses.Length; index++)
                {
                    var address = ProbeAddresses[index];
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? cell = null;
                    try
                    {
                        cell = source.Range[address];
                        var directPrecedents = Read(() => cell.DirectPrecedents);
                        var directDependents = Read(() => cell.DirectDependents);
                        var precedents = Read(() => cell.Precedents);
                        AssertRelation(DirectPrecedents[index], directPrecedents);
                        AssertRelation(DirectDependents[index], directDependents);
                        AssertRelation(AllPrecedents[index], precedents);
                        output.WriteLine($"active {address}: direct-precedents={directPrecedents}; " +
                            $"direct-dependents={directDependents}; precedents={precedents}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
                other.Activate();
                Excel.Range? inactive = null;
                try
                {
                    inactive = source.Range["C1"];
                    var relation = Read(() => inactive.DirectPrecedents);
                    AssertRelation("$B$1", relation);
                    output.WriteLine($"inactive C1: direct-precedents={relation}");
                }
                finally
                {
                    ComUtilities.Release(ref inactive);
                }
            }
            finally
            {
                try
                {
                    originalSheet?.Activate();
                    originalSelection?.Select();
                }
                finally
                {
                    ComUtilities.Release(ref other);
                    ComUtilities.Release(ref source);
                    ComUtilities.Release(ref originalSelection);
                    ComUtilities.Release(ref originalSheet);
                }
            }
        });
    }

    private static void AssertRelation(string? expected, (string? Address, int? ErrorCode, string? Error) actual)
    {
        Assert.Equal(expected, actual.Address);
        if (expected is null)
        {
            Assert.Equal(unchecked((int)0x800A03EC), actual.ErrorCode);
            Assert.False(string.IsNullOrWhiteSpace(actual.Error));
        }
        else
        {
            Assert.Null(actual.ErrorCode);
            Assert.Null(actual.Error);
        }
    }

    private static (string? Address, int? ErrorCode, string? Error) Read(Func<Excel.Range> getter)
    {
        Excel.Range? related = null;
        try
        {
            related = getter();
            return (related.Address, null, null);
        }
        catch (COMException exception) when (exception.HResult == unchecked((int)0x800A03EC))
        {
            return (null, exception.HResult, exception.Message);
        }
        finally
        {
            ComUtilities.Release(ref related);
        }
    }
}
