using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Slicer")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceSlicerStyleLayoutTests(PersistentServiceDataModelFixture fixture, ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceDataModelFixture>
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ApplyStyle_TwoExistingMulticolumnSlicersKeepLayoutAndFilters(bool model)
    {
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Office.LanguageSettings? language = null;
            try
            {
                language = ctx.App.LanguageSettings;
                output.WriteLine($"Excel UI language: {language.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI]}; " +
                    $"Excel country setting: {ctx.App.International[Excel.XlApplicationInternational.xlCountrySetting]}; " +
                    $"Excel decimal separator: {ctx.App.DecimalSeparator}; system separators: {ctx.App.UseSystemSeparators}.");
            }
            finally { ComUtilities.Release(ref language); }
        });
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        const string pivot = "StylePivot";
        string field = model ? "[CustomersTable].[Region]" : "Region";
        if (model)
            _fixture.Send("pivottable.create-from-datamodel", new { tableName = "SalesTable", destinationSheet = sheet, destinationCell = "A1", pivotTableName = pivot });
        else
        {
            RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B4",
                [["Region", "Amount"], ["North", 10], ["South", 20], ["North", 30]]));
            _fixture.Send("pivottable.create-from-range", new { sourceSheet = sheet, sourceRange = "A1:B4", destinationSheet = sheet, destinationCell = "D1", pivotTableName = pivot });
        }
        _fixture.Send("pivottablefield.add-row-field", new { pivotTableName = pivot, fieldName = field });
        _fixture.Send("pivottablefield.add-value-field", new { pivotTableName = pivot, fieldName = model ? "[Measures].[Total Sales]" : "Amount" });
        foreach (var (name, position) in new[] { ("StyleFirst", "J2"), ("StyleSecond", "N2") })
            _fixture.Send("slicer.create-slicer", new { pivotTableName = pivot, fieldName = field, slicerName = name, destinationSheet = sheet, position });
        _fixture.Send("slicer.set-slicer-selection", new { slicerName = "StyleFirst", selectedItems = new List<string> { "North" } });
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.SlicerCaches? caches = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                for (int index = 1; index <= caches.Count; index++)
                {
                    Excel.SlicerCache? cache = null;
                    Excel.Slicers? slicers = null;
                    try
                    {
                        cache = caches[index];
                        slicers = cache.Slicers;
                        for (int item = 1; item <= slicers.Count; item++)
                        {
                            Excel.Slicer? slicer = null;
                            try
                            {
                                slicer = slicers[item];
                                if (slicer.Name is "StyleFirst" or "StyleSecond")
                                {
                                    slicer.NumberOfColumns = slicer.Name == "StyleFirst" ? 2 : 3;
                                    slicer.Width = 260;
                                    slicer.Caption = slicer.Name + " keep";
                                }
                            }
                            finally { ComUtilities.Release(ref slicer); }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref slicers);
                        ComUtilities.Release(ref cache);
                    }
                }
            }
            finally { ComUtilities.Release(ref caches); }
        });
        using var firstBefore = State("StyleFirst");
        using var secondBefore = State("StyleSecond");
        Assert.Equal(2, firstBefore.RootElement.GetProperty("slicer").GetProperty("columnCount").GetInt32());
        Assert.Equal(3, secondBefore.RootElement.GetProperty("slicer").GetProperty("columnCount").GetInt32());
        Assert.Equal(model ? "[CustomersTable].[Region].&[North]" : "North",
            Assert.Single(firstBefore.RootElement.GetProperty("slicer").GetProperty("selectedItems").EnumerateArray()).GetString());
        var dataBefore = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = pivot });
        string styleName = $"DashboardStyle_{Guid.NewGuid():N}";
        _fixture.Send("slicer.create-style", new { styleName, sourceStyleName = "SlicerStyleLight1", slicerStyleOptions = new { selectedItemWithData = new { fillColor = "#E8F0FE", borderColor = "#336699" } } });
        _fixture.Send("slicer.apply-style", new { slicerName = "StyleFirst", styleName });
        using var neighbor = State("StyleSecond");
        Assert.Equal(secondBefore.RootElement.GetProperty("slicer").GetRawText(), neighbor.RootElement.GetProperty("slicer").GetRawText());
        _fixture.Send("slicer.apply-style", new { slicerName = "StyleSecond", styleName = "SlicerStyleLight2" });
        using var firstAfter = State("StyleFirst");
        using var secondAfter = State("StyleSecond");
        PersistentServiceSlicerStyleTests.AssertStyleOnly(firstBefore, firstAfter, styleName);
        PersistentServiceSlicerStyleTests.AssertStyleOnly(secondBefore, secondAfter, "SlicerStyleLight2");
        Assert.Equal(dataBefore.Result, _fixture.Send("pivottablecalc.get-data", new { pivotTableName = pivot }).Result);
        _fixture.Send("slicer.set-slicer-selection", new { slicerName = "StyleFirst", selectedItems = new List<string> { "South" } });
        using var selected = State("StyleSecond");
        Assert.Equal(model ? "[CustomersTable].[Region].&[South]" : "South",
            Assert.Single(selected.RootElement.GetProperty("slicer").GetProperty("selectedItems").EnumerateArray()).GetString());
        Assert.NotEqual(dataBefore.Result, _fixture.Send("pivottablecalc.get-data", new { pivotTableName = pivot }).Result);
    }

    private JsonDocument State(string name)
    {
        var response = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        Assert.True(response.Success, response.ErrorMessage);
        return JsonDocument.Parse(response.Result!);
    }
}
