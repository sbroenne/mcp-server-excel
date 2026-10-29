using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    private static readonly string[] BusyRegressionFormatRanges = ["A1:D1", "A2:A5", "B2:B5", "C2:C5", "D2:D5", "F1:F5"];

    [Theory]
    [InlineData(false, "ColumnStacked")]
    [InlineData(true, "ColumnStacked")]
    [InlineData(false, "BarStacked")]
    [InlineData(true, "BarStacked")]
    public async Task RepeatedChartSourceWritesAndFormatting_PreserveSessionAndWorksheetOrder(bool attached, string chartType)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var owner = await CreateSessionAsync(service, Path.Join(directory, "chart-workflow.xlsx"), show: true);
            sessions[owner] = 0;
            await WriteMarkerAsync(service, owner, "Unchanged neighboring workbook");
            var ownerBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(owner));
            var session = owner;
            if (attached)
            {
                var name = ownerBatch.Execute((context, _) => AddReviewWorkbook(context.App));
                session = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name));
                sessions[session] = 0;
            }
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var autoSave = batch.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            Assert.False(autoSave);
            await CreateSheetAsync(service, session, "Inputs");
            await CreateSheetAsync(service, session, "Report");
            var order = ReadOrder();
            await SendAsync("range.set-values", new
            {
                sheetName = "Inputs",
                rangeAddress = "A1:D5",
                values = new object[][] { ["Category", "One", "Two", "Three"],
                    ["A", 1, 2, 3], ["B", 4, 5, 6], ["C", 7, 8, 9], ["D", 10, 11, 12] }
            });
            await SendAsync("chart.create-from-range", new
            {
                sheetName = "Inputs",
                sourceRangeAddress = "A1:D5",
                chartType,
                chartName = "BusyRegression"
            });
            for (var iteration = 1; iteration <= 20; iteration++)
            {
                await SendAsync("range.set-formulas", new
                {
                    sheetName = "Inputs",
                    rangeAddress = "B2:D5",
                    overwritePolicy = "allow",
                    formulas = new string[][] {
                        [$"={iteration}+1", $"={iteration}+2", $"={iteration}+3"],
                        [$"={iteration}+4", $"={iteration}+5", $"={iteration}+6"],
                        [$"={iteration}+7", $"={iteration}+8", $"={iteration}+9"],
                        [$"={iteration}+10", $"={iteration}+11", $"={iteration}+12"] }
                });
                await SendAsync("rangeformat.format", new
                {
                    sheetName = "Inputs",
                    rangeAddresses = BusyRegressionFormatRanges,
                    formatOptions = new { bold = true, fontColor = "#123456", fillColor = "#EEEEEE" }
                });
                await SendAsync("chartconfig.set-data-labels", new
                {
                    chartName = "BusyRegression",
                    showValue = true,
                    seriesIndex = 2
                });
                await SendAsync("chartconfig.set-data-labels", new
                {
                    chartName = "BusyRegression",
                    showValue = false,
                    showPercentage = false,
                    showSeriesName = false,
                    showCategoryName = false,
                    showBubbleSize = false
                });
            }

            Assert.Equal(order, ReadOrder());
            Assert.Same(batch, service.SessionManager.GetSession(session));
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal(autoSave, batch.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn)));
            batch.Execute((context, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.ChartObject? chartObject = null;
                Excel.Chart? chart = null;
                Excel.SeriesCollection? collection = null;
                Excel.Series? series = null;
                Excel.Range? cell = null;
                Excel.Font? font = null;
                try
                {
                    sheets = context.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets["Inputs"];
                    cell = sheet.Range["B2"];
                    Assert.Equal("=20+1", (object)cell.Formula);
                    Assert.Equal(21d, (object)cell.Value2);
                    font = cell.Font;
                    Assert.True(Assert.IsType<bool>((object)font.Bold));
                    chartObject = (Excel.ChartObject)sheet.ChartObjects("BusyRegression");
                    chart = chartObject.Chart;
                    collection = (Excel.SeriesCollection)chart.SeriesCollection();
                    Assert.Equal(3, collection.Count);
                    for (var index = 1; index <= 3; index++)
                    {
                        series = collection.Item(index);
                        Assert.False(series.HasDataLabels);
                        Assert.Equal(new[] { 20d + index, 23d + index, 26d + index, 29d + index },
                            Assert.IsAssignableFrom<Array>((object)series.Values).Cast<object>().Select(Convert.ToDouble));
                        ComUtilities.Release(ref series);
                    }
                    Assert.False(context.Book.Saved);
                }
                finally
                {
                    ComUtilities.Release(ref font);
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref series);
                    ComUtilities.Release(ref collection);
                    ComUtilities.Release(ref chart);
                    ComUtilities.Release(ref chartObject);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
            if (attached)
                Assert.Equal("Unchanged neighboring workbook", await ReadMarkerAsync(service, owner));

            async Task SendAsync(string command, object arguments) =>
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = session,
                    Args = JsonSerializer.Serialize(arguments, ServiceProtocol.JsonOptions)
                }));

            string[] ReadOrder() => batch.Execute((context, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                try
                {
                    sheets = context.Book.Worksheets;
                    var names = new string[sheets.Count];
                    for (var index = 1; index <= names.Length; index++)
                    {
                        sheet = (Excel.Worksheet)sheets[index];
                        names[index - 1] = sheet.Name;
                        ComUtilities.Release(ref sheet);
                    }
                    return names;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
        });
    }
}
