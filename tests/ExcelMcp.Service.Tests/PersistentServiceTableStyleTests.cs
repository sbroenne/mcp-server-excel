using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "TableStyles")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceTableStyleTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public async Task TableStyle_LifecyclePreservesDefinitionAcrossReopen()
    {
        string styleName = $"TableStyle_{Guid.NewGuid():N}";
        var created = _fixture.Send("workbook.create-table-style", new { styleName, sourceStyleName = "TableStyleMedium2" });
        Assert.True(created.Success, created.ErrorMessage);
        using (var read = JsonDocument.Parse(created.Result!))
        {
            Assert.False(read.RootElement.GetProperty("style").GetProperty("builtIn").GetBoolean());
            Assert.Equal(styleName, read.RootElement.GetProperty("style").GetProperty("name").GetString());
        }
        var changed = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new
            {
                showAsAvailableTableStyle = true,
                elements = new[] { new { elementType = "xlHeaderRow", bold = true, fillColor = "#123456", fontColor = "#FFFFFF" } }
            }
        });
        Assert.True(changed.Success, changed.ErrorMessage);
        await _fixture.SaveAndReopenAsync();
        using (var read = JsonDocument.Parse(_fixture.Send("workbook.get-table-style", new { styleName }).Result!))
        {
            var header = Element(read, "xlHeaderRow");
            Assert.True(header.GetProperty("hasFormat").GetBoolean());
            Assert.Equal("#123456", header.GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
        }
        Assert.True(_fixture.Send("workbook.delete-table-style", new { styleName }).Success);
        var missing = await _fixture.SendForFailureAsync("workbook.get-table-style", new { styleName });
        Assert.Equal("NotFound", missing.ErrorCategory);
    }

    [Fact]
    public void TableStyle_InspectionIncludesEveryDistinctNativeElementWithoutInventingUnsetFormats()
    {
        using var read = JsonDocument.Parse(_fixture.Send("workbook.get-table-style", new { styleName = "TableStyleMedium2" }).Result!);
        var elements = read.RootElement.GetProperty("style").GetProperty("elements").EnumerateArray().ToList();
        Assert.Equal(Enum.GetValues<Excel.XlTableStyleElementType>().Distinct().Count(), elements.Count);
        Assert.Equal(elements.Count, elements.Select(element => element.GetProperty("nativeType").GetInt32()).Distinct().Count());
        var empty = elements.First(element => !element.GetProperty("hasFormat").GetBoolean());
        Assert.Equal(JsonValueKind.Null, empty.GetProperty("font").ValueKind);
        Assert.Equal(JsonValueKind.Null, empty.GetProperty("fill").ValueKind);
        Assert.Equal(JsonValueKind.Null, empty.GetProperty("stripeSize").ValueKind);
        using var list = JsonDocument.Parse(_fixture.Send("workbook.list-table-styles", new { }).Result!);
        Assert.Contains(list.RootElement.GetProperty("styles").EnumerateArray(),
            style => style.GetProperty("name").GetString() == "TableStyleMedium2" && style.GetProperty("builtIn").GetBoolean());
    }

    [Theory]
    [InlineData("xlHeaderRow")]
    [InlineData("xlFirstHeaderCell")]
    [InlineData("xlSubtotalRow1")]
    [InlineData("xlSlicerSelectedItemWithData")]
    [InlineData("xlTimelineSelectedTimeBlock")]
    public void TableStyle_UpdatesPreviouslyUnsetAndFormattedElementsAndCanClearThem(string elementType)
    {
        string styleName = CreateStyle();
        var update = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new
            {
                elements = new[] { new { elementType, bold = false, italic = true, fillThemeColor = 6, fontThemeColor = 5, fillTintAndShade = 0.25 } }
            }
        });
        using (var read = JsonDocument.Parse(update.Result!))
        {
            var element = Element(read, elementType);
            Assert.True(element.GetProperty("hasFormat").GetBoolean());
            Assert.False(element.GetProperty("font").GetProperty("bold").GetBoolean());
            Assert.True(element.GetProperty("font").GetProperty("italic").GetBoolean());
            Assert.Equal(6, element.GetProperty("fill").GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.InRange(element.GetProperty("fill").GetProperty("color").GetProperty("tintAndShade").GetDouble(), 0.24995, 0.25005);
        }
        var cleared = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType, clear = true } } }
        });
        using var result = JsonDocument.Parse(cleared.Result!);
        Assert.False(Element(result, elementType).GetProperty("hasFormat").GetBoolean());
    }

    [Theory]
    [InlineData("xlRowStripe1")]
    [InlineData("xlRowStripe2")]
    [InlineData("xlColumnStripe1")]
    [InlineData("xlColumnStripe2")]
    public async Task TableStyle_StripeSizesAreStoredAndPersist(string elementType)
    {
        string styleName = CreateStyle();
        var updated = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType, fillColor = "#123456", stripeSize = 3 } } }
        });
        using (var result = JsonDocument.Parse(updated.Result!))
            Assert.Equal(3, Element(result, elementType).GetProperty("stripeSize").GetInt32());
        await _fixture.SaveAndReopenAsync();
        using var reopened = JsonDocument.Parse(_fixture.Send("workbook.get-table-style", new { styleName }).Result!);
        Assert.Equal(3, Element(reopened, elementType).GetProperty("stripeSize").GetInt32());
    }

    [Theory]
    [InlineData("Left", "xlEdgeLeft")]
    [InlineData("Top", "xlEdgeTop")]
    [InlineData("Bottom", "xlEdgeBottom")]
    [InlineData("Right", "xlEdgeRight")]
    [InlineData("InsideHorizontal", "xlInsideHorizontal")]
    [InlineData("InsideVertical", "xlInsideVertical")]
    public void TableStyle_SixNativeBordersAreUpdated(string position, string edge)
    {
        string styleName = CreateStyle();
        var update = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new
            {
                elements = new[] { new { elementType = "xlWholeTable", borders = new[] { new { position, lineStyle = "dash", color = "#123456" } } } }
            }
        });
        using var result = JsonDocument.Parse(update.Result!);
        var borders = Element(result, "xlWholeTable").GetProperty("borders").EnumerateArray();
        var border = Assert.Single(borders, item => item.GetProperty("edge").GetString() == edge);
        Assert.Equal(-4115, border.GetProperty("lineStyle").GetInt32());
        Assert.Equal("#123456", border.GetProperty("color").GetProperty("rgb").GetString());
    }

    [Theory]
    [InlineData("showAsAvailableTableStyle")]
    [InlineData("showAsAvailablePivotTableStyle")]
    [InlineData("showAsAvailableSlicerStyle")]
    [InlineData("showAsAvailableTimelineStyle")]
    public void TableStyle_AvailabilityChangesPreserveOmittedFlags(string flag)
    {
        string styleName = CreateStyle();
        Assert.True(_fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new Dictionary<string, object?> { [flag] = false }
        }).Success);
        var changed = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType = "xlHeaderRow", bold = false } } }
        });
        using var read = JsonDocument.Parse(changed.Result!);
        Assert.False(read.RootElement.GetProperty("style").GetProperty(flag).GetBoolean());
    }

    [Theory]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","bold":false},{"elementType":"xlRowStripe1","stripeSize":0}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","bold":false},{"elementType":"xlRowStripe2","stripeSize":3}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","bold":false},{"elementType":"xlHeaderRow","italic":true}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","clear":true,"bold":false}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","stripeSize":2}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","fontName":"Arial"}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","fontSize":20}]}""")]
    [InlineData("""{"elements":[{"elementType":"xlHeaderRow","borders":[{"position":"DiagonalUp","lineStyle":"dash"}]}]}""")]
    [InlineData("""{"elements":[{"elementType":"99","bold":true}]}""")]
    [InlineData("""{"elements":[null]}""")]
    public async Task TableStyle_InvalidOptionsPreserveEntireDefinition(string options)
    {
        string styleName = CreateStyle();
        var before = _fixture.Send("workbook.get-table-style", new { styleName });
        using var parsed = JsonDocument.Parse(options);
        var rejected = await _fixture.SendForFailureAsync("workbook.update-table-style", new { styleName, tableStyleOptions = parsed.RootElement });
        Assert.False(rejected.Success);
        Assert.Equal(before.Result, _fixture.Send("workbook.get-table-style", new { styleName }).Result);
    }

    [Theory]
    [InlineData("update-table-style")]
    [InlineData("delete-table-style")]
    public async Task TableStyle_BuiltInMutationsPreserveDefinition(string action)
    {
        string styleName = "TableStyleMedium2";
        var before = _fixture.Send("workbook.get-table-style", new { styleName });
        object args = action == "update-table-style"
            ? new { styleName, tableStyleOptions = new { elements = new[] { new { elementType = "xlHeaderRow", bold = false } } } }
            : new { styleName };
        var rejected = await _fixture.SendForFailureAsync($"workbook.{action}", args);
        Assert.False(rejected.Success);
        Assert.Contains("Built-in", rejected.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before.Result, _fixture.Send("workbook.get-table-style", new { styleName }).Result);
    }

    [Fact]
    public async Task TableStyle_MissingAndDuplicateNamesPreserveCatalogue()
    {
        var before = _fixture.Send("workbook.list-table-styles", new { });
        foreach (var args in new[]
        {
                new { styleName = "TableStyleMedium2", sourceStyleName = "TableStyleMedium2" },
                new { styleName = $"Missing_{Guid.NewGuid():N}", sourceStyleName = "MissingStyle" }
            })
        {
            var failed = await _fixture.SendForFailureAsync("workbook.create-table-style", args);
            Assert.False(failed.Success);
        }
        Assert.Equal(before.Result, _fixture.Send("workbook.list-table-styles", new { }).Result);
    }

    private string CreateStyle()
    {
        string name = $"Style_{Guid.NewGuid():N}";
        var result = _fixture.Send("workbook.create-table-style", new { styleName = name, sourceStyleName = "TableStyleMedium2" });
        Assert.True(result.Success, result.ErrorMessage);
        return name;
    }

    [Theory]
    [InlineData("""{"bold":true}""", "font", "bold")]
    [InlineData("""{"italic":true}""", "font", "italic")]
    [InlineData("""{"strikethrough":true}""", "font", "strikethrough")]
    public void TableStyle_PreviouslyUnsetFontComponentsPersistWithoutFill(string options, string component, string property)
    {
        string styleName = CreateStyle();
        using var parsed = JsonDocument.Parse(options);
        var element = JsonSerializer.Deserialize<Dictionary<string, object>>(parsed.RootElement.GetRawText())!;
        element["elementType"] = "xlFirstHeaderCell";
        var response = _fixture.Send("workbook.update-table-style", new { styleName, tableStyleOptions = new { elements = new[] { element } } });
        Assert.True(response.Success, response.ErrorMessage);
        using var read = JsonDocument.Parse(response.Result!);
        Assert.True(Element(read, "xlFirstHeaderCell").GetProperty(component).GetProperty(property).GetBoolean());
    }

    private static JsonElement Element(JsonDocument document, string elementType) =>
        Assert.Single(document.RootElement.GetProperty("style").GetProperty("elements").EnumerateArray(),
            element => element.GetProperty("nativeType").GetInt32() == (int)Enum.Parse<Excel.XlTableStyleElementType>(elementType));

    [Theory]
    [InlineData("Name")]
    [InlineData("Size")]
    [InlineData("Subscript")]
    [InlineData("Superscript")]
    public void NativeTableStyleElements_RejectUnsupportedFontProperties(string property)
    {
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            Excel.TableStyle? template = null;
            Excel.TableStyleElements? elements = null;
            Excel.TableStyleElement? element = null;
            Excel.Font? font = null;
            try
            {
                styles = ctx.Book.TableStyles;
                string name = $"Probe_{Guid.NewGuid():N}";
                template = styles["TableStyleMedium2"];
                style = template.Duplicate(name);
                ComUtilities.Release(ref style);
                style = styles[name];
                elements = style.TableStyleElements;
                element = elements.Item(Excel.XlTableStyleElementType.xlHeaderRow);
                font = element.Font;
                var error = Record.Exception(() =>
                {
                    switch (property)
                    {
                        case "Name": font.Name = "Arial"; break;
                        case "Size": font.Size = 15d; break;
                        case "Subscript": font.Subscript = true; break;
                        case "Superscript": font.Superscript = true; break;
                    }
                });
                Assert.True(error is System.Runtime.InteropServices.COMException or ArgumentException, error?.ToString());
                style.Delete();
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref element);
                ComUtilities.Release(ref elements);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref template);
                ComUtilities.Release(ref styles);
            }
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TableStyle_UpdatesExistingTableAndPivotUsersWithoutReapplying(bool pivot)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string name = $"Object_{Guid.NewGuid():N}";
        string styleName = $"Shared_{Guid.NewGuid():N}";
        Assert.True(_fixture.Send("range.set-values", new
        {
            sheetName,
            rangeAddress = "A1:B3",
            values = new object[][] { ["Region", "Sales"], ["North", 10], ["South", 20] }
        }).Success);
        Assert.True(_fixture.Send("workbook.create-table-style", new
        {
            styleName,
            sourceStyleName = pivot ? "PivotStyleMedium9" : "TableStyleMedium2"
        }).Success);
        if (pivot)
        {
            Assert.True(_fixture.Send("pivottable.create-from-range", new
            {
                sourceSheet = sheetName,
                sourceRange = "A1:B3",
                destinationSheet = sheetName,
                destinationCell = "D1",
                pivotTableName = name
            }).Success);
            Assert.True(_fixture.Send("pivottablefield.add-row-field", new { pivotTableName = name, fieldName = "Region" }).Success);
            Assert.True(_fixture.Send("pivottablefield.add-value-field", new { pivotTableName = name, fieldName = "Sales", aggregationFunction = "sum" }).Success);
            Assert.True(_fixture.Send("pivottablecalc.set-layout-options", new { pivotTableName = name, layoutOptions = new { styleName } }).Success);
        }
        else
            Assert.True(_fixture.Send("table.create", new { sheetName, tableName = name, rangeAddress = "A1:B3", tableStyle = styleName }).Success);
        Assert.True(_fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType = "xlHeaderRow", fillColor = "#123456", bold = false } } }
        }).Success);
        await _fixture.SaveAndReopenAsync();
        var read = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = pivot ? "D1" : "A1", view = "displayed" });
        Assert.True(read.Success, read.ErrorMessage);
        using (var result = JsonDocument.Parse(read.Result!))
            Assert.Equal("#123456", result.RootElement.GetProperty("cells")[0].GetProperty("displayed").GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
        Assert.True(_fixture.Send("workbook.delete-table-style", new { styleName }).Success);
        var after = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = pivot ? "D1" : "A1", view = "displayed" });
        Assert.True(after.Success, after.ErrorMessage);
        using var final = JsonDocument.Parse(after.Result!);
        Assert.NotEqual("#123456", final.RootElement.GetProperty("cells")[0].GetProperty("displayed").GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
    }
}
