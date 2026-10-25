using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedCache_Create_ReusesNativeCacheWithoutCreatingAnother(bool table)
    {
        var batch = _fixture.BatchToken;
        var firstSheet = _fixture.CreateTestSheet(batch);
        var secondSheet = _fixture.CreateTestSheet(batch);
        var thirdSheet = _fixture.CreateTestSheet(batch);
        if (table)
            RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));

        RequireSuccess(CreateCacheReport(table, firstSheet, "First"));
        var firstIndex = ReadNativePivot(firstSheet, "First", pivot => pivot.CacheIndex);
        var cacheCount = CountNativePivotCaches();
        RequireSuccess(CreateCacheReport(table, secondSheet, "Second", "First", firstSheet));

        Assert.Equal(firstIndex, ReadNativePivot(secondSheet, "Second", pivot => pivot.CacheIndex));
        Assert.Equal(cacheCount, CountNativePivotCaches());
        RequireSuccess(CreateCacheReport(table, thirdSheet, "Independent"));
        Assert.NotEqual(firstIndex, ReadNativePivot(thirdSheet, "Independent", pivot => pivot.CacheIndex));
        Assert.Equal(cacheCount + 1, CountNativePivotCaches());
        foreach (var (sheet, name) in new[] { (firstSheet, "First"), (secondSheet, "Second"), (thirdSheet, "Independent") })
        {
            var nativeIndex = ReadNativePivot(sheet, name, pivot => pivot.CacheIndex);
            Assert.True(nativeIndex > 0);
            var read = RequireSuccess(_pivotCommands.Read(batch, name, sheet));
            Assert.Equal(nativeIndex, read.PivotTable.CacheIndex);
            var listed = Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables, item => item.Name == name);
            Assert.Equal(nativeIndex, listed.CacheIndex);
            using var json = System.Text.Json.JsonDocument.Parse(
                System.Text.Json.JsonSerializer.Serialize(read, ServiceProtocol.JsonOptions));
            Assert.Equal(nativeIndex, json.RootElement.GetProperty("pivotTable").GetProperty("cacheIndex").GetInt32());
        }

        foreach (var name in new[] { "First", "Second", "Independent" })
        {
            RequireSuccess(_pivotCommands.AddRowField(batch, name, name == "Second" ? "Product" : "Region"));
            RequireSuccess(_pivotCommands.AddValueField(batch, name, "Sales"));
            RequireSuccess(_pivotCommands.Refresh(batch, name));
            if (name == "Second")
                AssertCacheReportTotals(name, new Dictionary<string, double> { ["Gadget"] = 275, ["Widget"] = 375 }, 650);
            else
                AssertPivotSales(325, 325, name);
        }

        RequireSuccess(_pivotCommands.CreateSlicer(batch, "First", "Region", "SharedRegion", firstSheet, "H2"));
        RequireSuccess(_pivotCommands.ConnectPivotTable(batch, "SharedRegion", "Second"));
        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "SharedRegion", ["North"]));
        AssertFilteredPivot("First", "North", 325);
        AssertCacheReportTotals("Second", new Dictionary<string, double> { ["Gadget"] = 75, ["Widget"] = 250 }, 325);
        AssertPivotSales(325, 325, "Independent");

        RequireSuccess(_pivotCommands.SetSlicerSelection(batch, "SharedRegion", []));
        AssertPivotSales(325, 325, "First");
        AssertCacheReportTotals("Second", new Dictionary<string, double> { ["Gadget"] = 275, ["Widget"] = 375 }, 650);
        AssertPivotSales(325, 325, "Independent");

        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "First", "DoubleSales", "=Sales*2"));
        var calculated = Assert.Single(RequireSuccess(_pivotCommands.ListCalculatedFields(batch, "Second")).CalculatedFields);
        Assert.Equal("DoubleSales", calculated.Name);
        Assert.Equal("=Sales*2", calculated.Formula);
        AssertCalculatedValues("First", "DoubleSales",
            new Dictionary<string, double> { ["North"] = 650, ["South"] = 650 }, 1300);
        AssertCalculatedValues("Second", "DoubleSales",
            new Dictionary<string, double> { ["Gadget"] = 550, ["Widget"] = 750 }, 1300);
        Assert.Empty(RequireSuccess(_pivotCommands.ListCalculatedFields(batch, "Independent")).CalculatedFields);
        AssertPivotSales(325, 325, "Independent");
        AssertOriginalSales();
    }

    private void AssertCacheReportTotals(string name, Dictionary<string, double> expected, double total)
    {
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, name)).Values;
        Assert.Equal(expected.Count + 2, values.Count);
        Assert.All(values, row => Assert.Equal(2, row.Count));
        foreach (var pair in expected)
        {
            var row = Assert.Single(values, row => string.Equals(row[0]?.ToString(), pair.Key, StringComparison.Ordinal));
            Assert.Equal(pair.Value, Convert.ToDouble(row[1], CultureInfo.InvariantCulture));
        }
        Assert.Equal(total, Convert.ToDouble(values[^1][1], CultureInfo.InvariantCulture));
    }

    private PivotTableCreateResult CreateCacheReport(bool table, string sheet, string name,
        string? shareWith = null, string? shareSheet = null) =>
        table
            ? RequireSuccess(_pivotCommands.CreateFromTable(_fixture.BatchToken, "SharedSales", sheet, "A1", name,
                shareWith, shareSheet))
            : RequireSuccess(_pivotCommands.CreateFromRange(_fixture.BatchToken, _salesSheetName, "$A$1:$D$6", sheet, "A1", name,
                shareWith, shareSheet));

    private int CountNativePivotCaches() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.PivotCaches? caches = null;
            try
            {
                caches = context.Book.PivotCaches();
                return caches.Count;
            }
            finally
            {
                ComUtilities.Release(ref caches);
            }
        });

    [Theory]
    [InlineData(false, "First", null)]
    [InlineData(true, "First", null)]
    [InlineData(false, null, "Destination")]
    [InlineData(true, null, "Destination")]
    [InlineData(false, "", "Destination")]
    [InlineData(true, " ", "Destination")]
    [InlineData(false, "First", "")]
    [InlineData(true, "First", " ")]
    [InlineData(false, "", "")]
    [InlineData(true, " ", " ")]
    public void SharedCache_InvalidOptionPair_PreservesReportsAndCaches(bool table, string? name, string? sheet)
    {
        var batch = _fixture.BatchToken;
        var destination = _fixture.CreateTestSheet(batch);
        if (table)
            RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));
        RequireSuccess(CreateCacheReport(table, destination, "First"));
        var before = SnapshotPivot("First");
        var count = CountNativePivotCaches();

        var error = Assert.Throws<ArgumentException>(() =>
            CreateCacheReport(table, destination, "Rejected", name, sheet));

        Assert.Contains("shareCacheWith", error.Message, StringComparison.Ordinal);
        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(before, SnapshotPivot("First"));
        Assert.Equal("First", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SharedCache_MissingExactTarget_DoesNotCreateReportOrCache(bool table, bool missingSheet)
    {
        var batch = _fixture.BatchToken;
        var firstSheet = _fixture.CreateTestSheet(batch);
        var otherSheet = _fixture.CreateTestSheet(batch);
        if (table)
            RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));
        RequireSuccess(CreateCacheReport(table, firstSheet, "First"));
        var count = CountNativePivotCaches();
        var before = SnapshotPivot("First");

        Assert.Throws<InvalidOperationException>(() =>
            CreateCacheReport(table, otherSheet, "Rejected", "First", missingSheet ? "MissingSheet" : otherSheet));

        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(before, SnapshotPivot("First"));
        Assert.Equal("First", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData("different-sheet")]
    [InlineData("different-range")]
    [InlineData("table-to-range")]
    [InlineData("range-to-table")]
    [InlineData("different-table")]
    [InlineData("table-name-through-range")]
    [InlineData("structured-table-through-range")]
    public void SharedCache_SourceMismatch_DoesNotCreateReportOrCache(string scenario)
    {
        var batch = _fixture.BatchToken;
        var destination = _fixture.CreateTestSheet(batch);
        var otherSource = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, otherSource, "A1:D6",
            RequireSuccess(_commands.GetValues(batch, _salesSheetName, "A1:D6")).Values));
        RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));
        RequireSuccess(_tableCommands.Create(batch, otherSource, "OtherSales", "A1:D6", true));
        RequireSuccess(CreateCacheReport(scenario is "table-to-range" or "different-table", destination, "First"));
        var count = CountNativePivotCaches();
        var before = SnapshotPivot("First");

        var error = Assert.Throws<ArgumentException>(() =>
        {
            if (scenario is "range-to-table" or "different-table")
                _pivotCommands.CreateFromTable(batch, scenario == "different-table" ? "OtherSales" : "SharedSales",
                    destination, "H1", "Rejected", "First", destination);
            else
                _pivotCommands.CreateFromRange(batch, scenario == "different-sheet" ? otherSource : _salesSheetName,
                    scenario == "different-range" ? "A1:D5" :
                        scenario == "table-name-through-range" ? "SharedSales" :
                        scenario == "structured-table-through-range" ? "SharedSales[#All]" : "A1:D6",
                    destination, "H1", "Rejected", "First", destination);
        });

        Assert.Contains("source", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(before, SnapshotPivot("First"));
        Assert.Equal("First", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SharedCache_NativeSourceForms_ResolveQuotedSheetsAndWorkbookIdentity(bool table, bool qualified)
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = $"O'Brien {Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sourceSheet);
        RequireSuccess(_commands.SetValues(batch, sourceSheet, "A1:D6",
            RequireSuccess(_commands.GetValues(batch, _salesSheetName, "A1:D6")).Values));
        var destination = _fixture.CreateTestSheet(batch);
        RequireSuccess(_tableCommands.Create(batch, sourceSheet, "SharedSales", "A1:D6", true));
        RequireSuccess(table
            ? _pivotCommands.CreateFromTable(batch, "SharedSales", destination, "A1", "First")
            : _pivotCommands.CreateFromRange(batch, sourceSheet, "A1:D6", destination, "A1", "First"));
        var quotedSheet = sourceSheet.Replace("'", "''", StringComparison.Ordinal);
        ReadNativePivot(destination, "First", pivot =>
        {
            Excel.PivotCache? cache = null;
            try
            {
                cache = pivot.PivotCache();
                Assert.Equal(table ? "SharedSales" : $"'{quotedSheet}'!R1C1:R6C4", ((object)cache.SourceData));
                if (table && qualified)
                    cache.SourceData = $"'{quotedSheet}'!SharedSales";
                if (!table && qualified)
                    cache.SourceData = $"'[{System.IO.Path.GetFileName(_fixture.WorkbookPath)}]{quotedSheet}'!R1C1:R6C4";
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref cache);
            }
        });
        var count = CountNativePivotCaches();
        RequireSuccess(table
            ? _pivotCommands.CreateFromTable(batch, "SharedSales", destination, "H1", "Second", "First", destination)
            : _pivotCommands.CreateFromRange(batch, sourceSheet, "$A$1:$D$6", destination, "H1", "Second", "First", destination));
        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(ReadNativePivot(destination, "First", pivot => pivot.CacheIndex),
            ReadNativePivot(destination, "Second", pivot => pivot.CacheIndex));
        RequireSuccess(_pivotCommands.AddRowField(batch, "Second", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "Second", "Sales"));
        AssertPivotSales(325, 325, "Second");
        AssertOriginalSales();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedCache_RepeatedPivotNames_SelectsExactWorksheet(bool table)
    {
        var batch = _fixture.BatchToken;
        var decoySheet = _fixture.CreateTestSheet(batch);
        var selectedSheet = _fixture.CreateTestSheet(batch);
        var newSheet = _fixture.CreateTestSheet(batch);
        if (table)
            RequireSuccess(_tableCommands.Create(batch, _salesSheetName, "SharedSales", "A1:D6", true));
        RequireSuccess(_pivotCommands.CreateFromRange(batch, _salesSheetName, "A1:C6", decoySheet, "A1", "Repeated"));
        RequireSuccess(CreateCacheReport(table, selectedSheet, "Repeated"));
        var count = CountNativePivotCaches();
        var expected = ReadNativePivot(selectedSheet, "Repeated", pivot => pivot.CacheIndex);
        var decoy = ReadNativePivot(decoySheet, "Repeated", pivot => pivot.CacheIndex);
        Assert.NotEqual(decoy, expected);

        RequireSuccess(CreateCacheReport(table, newSheet, "Second", "Repeated", selectedSheet));

        Assert.Equal(expected, ReadNativePivot(newSheet, "Second", pivot => pivot.CacheIndex));
        Assert.Equal(decoy, ReadNativePivot(decoySheet, "Repeated", pivot => pivot.CacheIndex));
        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(3, RequireSuccess(_pivotCommands.List(batch)).PivotTables.Count);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData("C1:C4", "$A:$D", "$C$1:$C$4")]
    [InlineData("R1:R6", "$1:$6", "$R$1:$R$6")]
    public void SharedCache_NativeSourceMatching_WholeAxesMatchWithoutMatchingAmbiguousA1Cells(
        string nativeAddress, string wholeAxisAddress, string ambiguousCellAddress)
    {
        var batch = _fixture.BatchToken;
        var destination = _fixture.CreateTestSheet(batch);
        RequireSuccess(CreateCacheReport(false, destination, "First"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "First", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "First", "Sales"));
        AssertPivotSales(325, 325, "First");
        var before = SnapshotPivot("First");
        var count = CountNativePivotCaches();
        var firstIndex = ReadNativePivot(destination, "First", pivot => pivot.CacheIndex);
        var sourceBefore = ReadNativeCacheSource(destination, "First");
        Assert.True(firstIndex > 0);
        Assert.NotEmpty(sourceBefore);

        // Native source-reference matching is an internal contract Service cannot expose cheaply:
        // assigning whole-axis cache sources would materialize full-sheet data or invalid headers.
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sourceSheet = null;
            Excel.Range? wholeAxisSource = null;
            Excel.Range? ambiguousCellSource = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(context.Book, _salesSheetName);
                Assert.NotNull(sourceSheet);
                wholeAxisSource = sourceSheet.Range[wholeAxisAddress];
                ambiguousCellSource = sourceSheet.Range[nativeAddress];
                Assert.Equal(wholeAxisAddress, Convert.ToString(
                    (object)context.App.ConvertFormula(nativeAddress, Excel.XlReferenceStyle.xlR1C1,
                        Excel.XlReferenceStyle.xlA1, Excel.XlReferenceType.xlAbsolute),
                    CultureInfo.InvariantCulture));
                Assert.Equal(wholeAxisAddress,
                    wholeAxisSource.Address[true, true, Excel.XlReferenceStyle.xlA1]);
                Assert.Equal(ambiguousCellAddress,
                    ambiguousCellSource.Address[true, true, Excel.XlReferenceStyle.xlA1]);

                var reference = $"'{_salesSheetName}'!{nativeAddress}";
                Assert.True(PivotTableCommands.MatchesCreationSource(context.Book, context.App,
                    reference, sourceSheet, wholeAxisSource, null));
                Assert.False(PivotTableCommands.MatchesCreationSource(context.Book, context.App,
                    reference, sourceSheet, ambiguousCellSource, null));
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref ambiguousCellSource);
                ComUtilities.Release(ref wholeAxisSource);
                ComUtilities.Release(ref sourceSheet);
            }
        });

        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(before, SnapshotPivot("First"));
        Assert.Equal(firstIndex, ReadNativePivot(destination, "First", pivot => pivot.CacheIndex));
        Assert.Equal(sourceBefore, ReadNativeCacheSource(destination, "First"));
        Assert.Equal("First", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertPivotSales(325, 325, "First");
        AssertOriginalSales();
    }

    private string ReadNativeCacheSource(string sheet, string name) =>
        ReadNativePivot(sheet, name, pivot =>
        {
            Excel.PivotCache? cache = null;
            try
            {
                cache = pivot.PivotCache();
                return Assert.IsType<string>(((object)cache.SourceData));
            }
            finally
            {
                ComUtilities.Release(ref cache);
            }
        });

    [Theory]
    [InlineData("Reports]", false)]
    [InlineData("Reports]", true)]
    [InlineData("Reports][", false)]
    [InlineData("Reports][", true)]
    public void SharedCache_SourceWorksheetMatching_BracketedDirectoryRequiresExactIdentity(
        string directoryName, bool quotedSheet)
    {
        var sheetName = $"{(quotedSheet ? "O'Brien " : "Plain_")}{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(_fixture.BatchToken, sheetName);

        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? sourceRange = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                sourceRange = sheet.Range["A1:D6"];
                Assert.Equal("$A$1:$D$6",
                    sourceRange.Address[true, true, Excel.XlReferenceStyle.xlA1]);
                var workbookName = context.Book.Name;
                var workbookFullName = context.Book.FullName;
                var directory = System.IO.Path.GetDirectoryName(workbookFullName)!;
                var nativeQualifier = System.IO.Path.Combine(directory, $"[{workbookName}]{sheetName}");
                Assert.True(PivotTableCommands.MatchesCreationSource(context.Book, context.App,
                    $"'{nativeQualifier.Replace("'", "''", StringComparison.Ordinal)}'!R1C1:R6C4",
                    sheet, sourceRange, null));

                // Excel rejects bracketed save paths, so test the qualifier contract with a synthetic path
                // while checking its native workbook/range caller against the actual saved identity.
                var bracketedDirectory = System.IO.Path.Combine(directory, directoryName);
                var bracketedFullName = System.IO.Path.Combine(bracketedDirectory, workbookName);
                var qualifier = System.IO.Path.Combine(bracketedDirectory, $"[{workbookName}]{sheetName}");
                Assert.False(PivotTableCommands.MatchesCreationSource(context.Book, context.App,
                    $"'{qualifier.Replace("'", "''", StringComparison.Ordinal)}'!R1C1:R6C4",
                    sheet, sourceRange, null));
                foreach (var rejected in new[]
                {
                    System.IO.Path.Combine(bracketedDirectory, "foreign", $"[{workbookName}]{sheetName}"),
                    System.IO.Path.Combine(bracketedDirectory, $"[Other.xlsx]{sheetName}"),
                    System.IO.Path.Combine(bracketedDirectory, $"[{workbookName}{sheetName}"),
                    $"{workbookName}]{sheetName}", $"[{workbookName}]OtherSheet"
                })
                {
                    Assert.False(PivotTableCommands.MatchesSourceWorksheet(workbookName, bracketedFullName,
                        $"'{rejected.Replace("'", "''", StringComparison.Ordinal)}'", sheetName));
                }
                Assert.True(PivotTableCommands.MatchesSourceWorksheet(workbookName, bracketedFullName,
                    $"'{qualifier.Replace("'", "''", StringComparison.Ordinal)}'", sheetName));
                if (!quotedSheet)
                    Assert.True(PivotTableCommands.MatchesSourceWorksheet(workbookName, bracketedFullName,
                        qualifier, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref sourceRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    [Theory]
    [InlineData("foreign-workbook-name")]
    [InlineData("foreign-workbook-path")]
    [InlineData("unresolvable-name")]
    public void SharedCache_ForeignOrUnresolvableSource_PreservesReportAndCache(string scenario)
    {
        var batch = _fixture.BatchToken;
        var destination = _fixture.CreateTestSheet(batch);
        RequireSuccess(CreateCacheReport(false, destination, "First"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "First", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "First", "Sales"));
        AssertPivotSales(325, 325, "First");
        var before = SnapshotPivot("First");
        var count = CountNativePivotCaches();
        var firstIndex = ReadNativePivot(destination, "First", pivot => pivot.CacheIndex);
        var sourceBefore = ReadNativeCacheSource(destination, "First");
        Assert.True(firstIndex > 0);
        Assert.NotEmpty(sourceBefore);

        // Excel rejects these references when assigned to SourceData, before the matcher runs.
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sourceSheet = null;
            Excel.Range? sourceRange = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(context.Book, _salesSheetName);
                Assert.NotNull(sourceSheet);
                sourceRange = sourceSheet.Range["A1:D6"];
                var reference = scenario switch
                {
                    "foreign-workbook-name" => $"'[OtherWorkbook.xlsx]{_salesSheetName}'!R1C1:R6C4",
                    "foreign-workbook-path" => $"'{System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(context.Book.FullName)!, "foreign",
                        $"[{context.Book.Name}]{_salesSheetName}")}'!R1C1:R6C4",
                    "unresolvable-name" => "UnresolvableSourceName",
                    _ => throw new ArgumentOutOfRangeException(nameof(scenario))
                };
                Assert.True(PivotTableCommands.MatchesCreationSource(context.Book, context.App,
                    sourceBefore, sourceSheet, sourceRange, null));
                Assert.False(PivotTableCommands.MatchesCreationSource(context.Book, context.App,
                    reference, sourceSheet, sourceRange, null));
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sourceRange);
                ComUtilities.Release(ref sourceSheet);
            }
        });

        Assert.Equal(count, CountNativePivotCaches());
        Assert.Equal(before, SnapshotPivot("First"));
        Assert.Equal(firstIndex, ReadNativePivot(destination, "First", pivot => pivot.CacheIndex));
        Assert.Equal(sourceBefore, ReadNativeCacheSource(destination, "First"));
        Assert.Equal("First", Assert.Single(RequireSuccess(_pivotCommands.List(batch)).PivotTables).Name);
        AssertPivotSales(325, 325, "First");
        AssertOriginalSales();
    }
}
