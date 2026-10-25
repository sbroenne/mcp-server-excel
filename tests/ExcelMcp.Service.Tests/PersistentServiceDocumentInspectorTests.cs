using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "DocumentInspector")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceDocumentInspectorTests : IAsyncLifetime, IDisposable
{
    private readonly PersistentServiceWorkbookFixture _fixture = new();
    private const string HiddenName = "InspectorHidden";
    private const string Seed = "Inspector test content";

    public Task InitializeAsync() => _fixture.InitializeAsync();
    public Task DisposeAsync() => _fixture.DisposeAsync();
    public void Dispose()
    {
        _fixture.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void NativeInspector_ReportsAndRemovesSeededHiddenSheetWithoutSaving()
    {
        SeedHiddenSheet();
        var selected = Assert.Single(ReadNativeInspectors(), item => item.Status == 1);
        var fixedStatus = _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Office.DocumentInspectors? inspectors = null;
            Office.DocumentInspector? inspector = null;
            try
            {
                inspectors = ctx.Book.DocumentInspectors;
                inspector = inspectors[selected.Index];
                Office.MsoDocInspectorStatus status;
                string text;
                inspector.Fix(out status, out text);
                Assert.NotEmpty(text);
                return (int)status;
            }
            finally
            {
                ComUtilities.Release(ref inspector);
                ComUtilities.Release(ref inspectors);
            }
        });
        Assert.Equal(0, fixedStatus);
        Assert.Equal(0, ReadNativeInspectors().Single(item => item.Index == selected.Index).Status);
        Assert.False(ReadState().HiddenExists);
        Assert.False(ReadState().Saved);
    }

    [Fact]
    public async Task InspectDocument_MatchesNativeResultsPreservesContentsAndReportsSavedState()
    {
        SeedHiddenSheet();
        SetHeader();
        foreach (bool saveFirst in new[] { false, true })
        {
            if (saveFirst)
                _fixture.ExecuteRawVerification((ctx, _) => ctx.Book.Save());
            var native = ReadNativeInspectors();
            bool expectedSavedAfter = ReadState().Saved;
            if (saveFirst)
                _fixture.ExecuteRawVerification((ctx, _) => ctx.Book.Save());
            var before = ReadState();
            Assert.Equal(saveFirst, before.Saved);
            Assert.Equal(Seed, before.Value);
            Assert.Equal(Excel.XlSheetVisibility.xlSheetHidden, before.Visibility);
            var bytes = ReadWorkbookBytes();
            var response = await _fixture.SendAsync("workbook.inspect-document", new { });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(before.Saved, result.RootElement.GetProperty("savedBefore").GetBoolean());
            Assert.Equal(expectedSavedAfter, result.RootElement.GetProperty("savedAfter").GetBoolean());
            Assert.Contains("not", result.RootElement.GetProperty("readLimitations").GetString(),
                StringComparison.OrdinalIgnoreCase);
            var items = result.RootElement.GetProperty("inspectors").EnumerateArray().ToArray();
            Assert.Equal(native.Count, items.Length);
            foreach (var expected in native)
            {
                var actual = Assert.Single(items, item => item.GetProperty("index").GetInt32() == expected.Index);
                Assert.Equal(expected.Name, actual.GetProperty("name").GetString());
                Assert.Equal(expected.Status, actual.GetProperty("statusCode").GetInt32());
                Assert.Equal(expected.Text, actual.GetProperty("results").GetString());
                Assert.Equal(expected.Status == 1 ? "issue-found" : "ok", actual.GetProperty("status").GetString());
            }
            Assert.Equal(2, items.Count(item => item.GetProperty("statusCode").GetInt32() == 1));
            Assert.Equal(before with { Saved = expectedSavedAfter }, ReadState());
            Assert.Equal(bytes, ReadWorkbookBytes());
        }
    }

    [Fact]
    public async Task FixDocumentInspector_RemovesOnlySelectedCategoryAndPersistsAfterExplicitSave()
    {
        SeedHiddenSheet();
        var selected = Assert.Single(ReadNativeInspectors(), item => item.Status == 1);
        SetHeader();
        var response = await _fixture.SendAsync("workbook.fix-document-inspector",
            new { inspectorIndex = selected.Index, inspectorName = selected.Name });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(selected.Index, result.RootElement.GetProperty("inspectorIndex").GetInt32());
        Assert.Equal(selected.Name, result.RootElement.GetProperty("inspectorName").GetString());
        Assert.Equal(0, result.RootElement.GetProperty("fix").GetProperty("statusCode").GetInt32());
        Assert.Equal(0, result.RootElement.GetProperty("inspection").GetProperty("statusCode").GetInt32());
        var state = ReadState();
        Assert.False(state.HiddenExists);
        Assert.False(state.Saved);
        Assert.Equal("Retained header", state.Header);
        Assert.Equal("=1+2", state.Formula);
        var remaining = Assert.Single(ReadNativeInspectors(), item => item.Status == 1);
        Assert.NotEqual(selected.Index, remaining.Index);
        await _fixture.SaveAndReopenAsync();
        Assert.False(ReadState().HiddenExists);
        Assert.Equal("Retained header", ReadState().Header);
        Assert.Equal("=1+2", ReadState().Formula);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(-1, false)]
    [InlineData(int.MaxValue, false)]
    [InlineData(1, true)]
    public async Task InvalidSelection_FailsBeforeChangingWorkbook(int index, bool mismatchedName)
    {
        SeedHiddenSheet();
        var selected = Assert.Single(ReadNativeInspectors(), item => item.Status == 1);
        var before = ReadState();
        var failure = await _fixture.SendForFailureAsync("workbook.fix-document-inspector",
            new
            {
                inspectorIndex = mismatchedName ? selected.Index : index,
                inspectorName = mismatchedName ? "Not the selected inspector" : selected.Name
            });
        Assert.Contains(mismatchedName ? "name" : "inspectorIndex", failure.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, ReadState());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnwritableWorkbook_FailsBeforeChangingWorkbook(bool readOnly)
    {
        SeedHiddenSheet();
        var selected = Assert.Single(ReadNativeInspectors(), item => item.Status == 1);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            if (readOnly)
            {
                ctx.Book.Save();
                ctx.Book.ChangeFileAccess(Excel.XlFileAccess.xlReadOnly);
            }
            else
                ctx.Book.Protect(Structure: true);
        });
        var inspection = await _fixture.SendAsync("workbook.inspect-document", new { });
        using var inspected = JsonDocument.Parse(inspection.Result!);
        Assert.True(inspected.RootElement.GetProperty("success").GetBoolean());
        var finding = Assert.Single(inspected.RootElement.GetProperty("inspectors").EnumerateArray(),
            item => item.GetProperty("index").GetInt32() == selected.Index);
        Assert.Equal(1, finding.GetProperty("statusCode").GetInt32());
        var before = ReadState();
        var failure = await _fixture.SendForFailureAsync("workbook.fix-document-inspector",
            new { inspectorIndex = selected.Index, inspectorName = selected.Name });
        Assert.Contains(readOnly ? "read-only" : "protected", failure.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, ReadState());
    }

    private void SeedHiddenSheet() => _fixture.ExecuteRawVerification((ctx, _) =>
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? hidden = null;
        Excel.Worksheet? visible = null;
        Excel.Range? cell = null;
        Excel.Range? formula = null;
        try
        {
            sheets = ctx.Book.Worksheets;
            visible = (Excel.Worksheet)sheets.Item[1];
            formula = visible.Range["A1"];
            formula.Formula = "=1+2";
            hidden = (Excel.Worksheet)sheets.Add();
            hidden.Name = HiddenName;
            cell = hidden.Range["A1"];
            cell.Value2 = Seed;
            hidden.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }
        finally
        {
            ComUtilities.Release(ref formula);
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref visible);
            ComUtilities.Release(ref hidden);
            ComUtilities.Release(ref sheets);
        }
    });

    private void SetHeader() => _fixture.ExecuteRawVerification((ctx, _) =>
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.PageSetup? setup = null;
        try
        {
            sheets = ctx.Book.Worksheets;
            sheet = (Excel.Worksheet)sheets.Item[sheets.Count];
            setup = sheet.PageSetup;
            setup.CenterHeader = "Retained header";
        }
        finally
        {
            ComUtilities.Release(ref setup);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    });

    private List<NativeInspector> ReadNativeInspectors() => _fixture.ExecuteRawVerification((ctx, _) =>
    {
        Office.DocumentInspectors? inspectors = null;
        List<NativeInspector> result = [];
        try
        {
            inspectors = ctx.Book.DocumentInspectors;
            int count = Convert.ToInt32(inspectors.Count);
            for (int index = 1; index <= count; index++)
            {
                Office.DocumentInspector? inspector = null;
                try
                {
                    inspector = inspectors[index];
                    Office.MsoDocInspectorStatus status;
                    string text;
                    inspector.Inspect(out status, out text);
                    result.Add(new NativeInspector(index, inspector.Name, (int)status, text));
                }
                finally
                {
                    ComUtilities.Release(ref inspector);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref inspectors);
        }
    });

    private WorkbookState ReadState() => _fixture.ExecuteRawVerification((ctx, _) =>
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? visible = null;
        Excel.Worksheet? hidden = null;
        Excel.Range? cell = null;
        Excel.Range? formula = null;
        Excel.PageSetup? setup = null;
        try
        {
            sheets = ctx.Book.Worksheets;
            visible = (Excel.Worksheet)sheets.Item[sheets.Count];
            hidden = ComUtilities.FindSheet(ctx.Book, HiddenName);
            formula = visible.Range["A1"];
            setup = visible.PageSetup;
            if (hidden is not null)
                cell = hidden.Range["A1"];
            return new WorkbookState(ctx.Book.Saved, hidden is not null,
                hidden?.Visible, ((object?)cell?.Value2)?.ToString(), setup.CenterHeader, ((object?)formula.Formula)?.ToString());
        }
        finally
        {
            ComUtilities.Release(ref setup);
            ComUtilities.Release(ref formula);
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref hidden);
            ComUtilities.Release(ref visible);
            ComUtilities.Release(ref sheets);
        }
    });

    private sealed record NativeInspector(int Index, string Name, int Status, string Text);
    private sealed record WorkbookState(bool Saved, bool HiddenExists,
        Excel.XlSheetVisibility? Visibility, string? Value, string Header, string? Formula);

    private byte[] ReadWorkbookBytes()
    {
        using var file = new FileStream(_fixture.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var copy = new MemoryStream();
        file.CopyTo(copy);
        return copy.ToArray();
    }
}
