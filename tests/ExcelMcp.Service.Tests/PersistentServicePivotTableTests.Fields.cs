using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for PivotTable field operations (Strategy Pattern: RegularPivotTableFieldStrategy).
/// Tests AddColumn, AddValue, AddFilter, Remove, Set* operations on Regular PivotTables.
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence.
/// Organized by category trait for Architecture Pattern clarity.
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    /// <inheritdoc/>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddColumnField_WithValidField_AddsFieldToColumns()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Act - No save needed
        var result = _pivotCommands.AddColumnField(batch, "TestPivot", "Product");

        // Assert
        RequireSuccess(result);
        Assert.Equal("Product", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
        RequireSuccess(result);
        AssertNativeField("TestPivot", "Product", PivotFieldArea.Column);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddValueField_WithValidField_AddsFieldToValues()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Act
        var result = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");

        // Assert
        RequireSuccess(result);
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
        RequireSuccess(result);
        AssertNativeField("TestPivot", "Sales", PivotFieldArea.Value, function: AggregationFunction.Sum);
        AssertScalarPivot(650);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddValueField_FromTableNumericColumn_AllowsSumAggregation()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(
            batch,
            _salesSheetName,
            "A1:D11",
            [
                ["Date", "Region", "Product", "Amount"],
                [new DateTime(2026, 1, 1), "North", "Widget", 100],
                [new DateTime(2026, 1, 2), "South", "Widget", 150],
                [new DateTime(2026, 1, 3), "North", "Gadget", 200],
                [new DateTime(2026, 1, 4), "West", "Widget", 75],
                [new DateTime(2026, 1, 5), "East", "Gadget", 125],
                [new DateTime(2026, 1, 6), "North", "Widget", 90],
                [new DateTime(2026, 1, 7), "South", "Gadget", 60],
                [new DateTime(2026, 1, 8), "West", "Widget", 110],
                [new DateTime(2026, 1, 9), "East", "Gadget", 80],
                [new DateTime(2026, 1, 10), "North", "Gadget", 130],
            ],
            overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_commands.SetNumberFormat(batch, _salesSheetName, "A2:A11", "m/d/yyyy"));
        RequireSuccess(_commands.SetNumberFormat(batch, _salesSheetName, "D2:D11", "$#,##0.00"));
        RequireSuccess(_tableCommands.Create(
            batch,
            _salesSheetName,
            "tblSales",
            "A1:D11",
            true,
            TableStylePresets.Medium2));

        var createResult = _pivotCommands.CreateFromTable(
            batch, "tblSales", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Act
        var fieldsResult = _pivotCommands.ListFields(batch, "TestPivot");
        var addResult = _pivotCommands.AddValueField(
            batch, "TestPivot", "Amount", AggregationFunction.Sum, "Total Amount");

        // Assert
        RequireSuccess(fieldsResult);
        var amountField = Assert.Single(fieldsResult.Fields, field => field.Name == "Amount");
        Assert.Equal("Number", amountField.DataType);

        RequireSuccess(addResult);
        Assert.Equal("Amount", addResult.FieldName);
        Assert.Equal(PivotFieldArea.Value, addResult.Area);
        Assert.Equal(AggregationFunction.Sum, addResult.Function);
        Assert.Equal("Number", addResult.DataType);
        RequireSuccess(addResult);
        AssertScalarPivot(1120);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddFilterField_WithValidField_AddsFieldToFilters()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Act
        var result = _pivotCommands.AddFilterField(batch, "TestPivot", "Region");

        // Assert
        RequireSuccess(result);
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Filter, result.Area);
        RequireSuccess(result);
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Filter);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetReportFilter_RegularPivot_IsRejectedWithoutChangingLayout()
    {
        var batch = _fixture.BatchToken;
        var created = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(created);
        var added = _pivotCommands.AddFilterField(batch, "TestPivot", "Region");
        RequireSuccess(added);

        var result = _pivotCommands.SetReportFilter(
            batch, _salesSheetName, "TestPivot", "Region", ["North"]);

        Assert.False(result.Success);
        Assert.Contains("not an OLAP/Data Model PivotTable", result.ErrorMessage);
        Assert.False(result.MayHavePartiallyChanged);
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Filter);
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void RemoveField_ExistingField_RemovesFromPivot()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add a field first
        var addResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        RequireSuccess(addResult);
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Row);

        // Act - Remove in same batch
        var result = _pivotCommands.RemoveField(batch, "TestPivot", "Region");

        // Assert
        RequireSuccess(result);

        // Verify field removed
        var infoResult = _pivotCommands.Read(batch, "TestPivot");
        RequireSuccess(infoResult);
        var regionField = infoResult.Fields.FirstOrDefault(f => f.Name == "Region");
        Assert.NotNull(regionField);
        Assert.Equal(PivotFieldArea.Hidden, regionField.Area);
        RequireSuccess(result);
        RequireSuccess(infoResult);
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Hidden);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFunction_ValueField_ChangesAggregation()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Sales as value field (default sum)
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        RequireSuccess(addResult);

        // Act - Change to Average in same batch
        AssertScalarPivot(650);
        var result = _pivotCommands.SetFieldFunction(batch, "TestPivot", "Sales", AggregationFunction.Average);

        // Assert
        RequireSuccess(result);
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(AggregationFunction.Average, result.Function);
        RequireSuccess(result);
        AssertNativeField("TestPivot", "Sales", PivotFieldArea.Value, function: AggregationFunction.Average);
        AssertScalarPivot(130);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldName_ExistingField_RenamesField()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        RequireSuccess(addResult);

        // Act
        var result = _pivotCommands.SetFieldName(batch, "TestPivot", "Sales", "Total Revenue");

        // Assert
        RequireSuccess(result);
        Assert.Equal("Total Revenue", result.CustomName);
        RequireSuccess(result);
        AssertNativeField("TestPivot", "Sales", PivotFieldArea.Value, caption: "Total Revenue");
        ReadNativePivot(_salesSheetName, "TestPivot", pivot =>
        {
            Microsoft.Office.Interop.Excel.PivotField? source = null;
            try
            {
                source = (Microsoft.Office.Interop.Excel.PivotField)pivot.PivotFields("Sales");
                Assert.Equal("Sales", source.Caption);
                return 0;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref source);
            }
        });
        AssertScalarPivot(650);
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    /// <summary>
    /// Verifies that SetFieldFormat with US currency format works correctly on any locale.
    /// The server should auto-translate format codes (number separators handled by UseSystemSeparators=false).
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFormat_USCurrencyFormat_RoundTripsCorrectly()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        RequireSuccess(addResult);

        // Act - Apply US currency format
        var result = _pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", "$#,##0.00");

        // Assert - Format should round-trip correctly (not corrupted by locale)
        RequireSuccess(result);
        AssertPivotNumberFormat("$#,##0.00", 650,
            _fixture.ExecuteRawVerification((ctx, _) => $"$650{ctx.FormatTranslator.DecimalSeparator}00"));
    }

    /// <summary>
    /// Verifies that SetFieldFormat with US date format works correctly on value fields.
    /// Uses a Count function on a date field to create a numeric value that can be formatted.
    /// The server auto-translates format codes (number separators handled by UseSystemSeparators=false).
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFormat_USPercentFormat_RoundTripsCorrectly()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        RequireSuccess(addResult);

        // Act - Apply US percent format (tests decimal separator preservation)
        var result = _pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", "0.00%");

        // Assert - Format should round-trip correctly (not corrupted by locale)
        RequireSuccess(result);
        AssertPivotNumberFormat("0.00%", 650,
            _fixture.ExecuteRawVerification((ctx, _) => $"65000{ctx.FormatTranslator.DecimalSeparator}00%"));
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFilter_RowField_AppliesFilter()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Region as row field
        var addResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        RequireSuccess(addResult);

        // Act
        var result = _pivotCommands.SetFieldFilter(batch, "TestPivot", "Region", ["North"]);

        // Assert
        RequireSuccess(result);
        Assert.Equal("Region", result.FieldName);
        Assert.NotEmpty(result.SelectedItems);
        RequireSuccess(result);
        Assert.Equal(["North"], result.SelectedItems);
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));
        RequireSuccess(_pivotCommands.Refresh(batch, "TestPivot"));
        var filtered = RequireSuccess(_pivotCommands.GetData(batch, "TestPivot"));
        Assert.Equal(3, filtered.Values.Count);
        Assert.Equal("North", filtered.Values[1][0]);
        Assert.Equal(325d, Convert.ToDouble(filtered.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(325d, Convert.ToDouble(filtered.Values[^1][^1], System.Globalization.CultureInfo.InvariantCulture));
        AssertOriginalSales();
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SortField_RowField_SortsData()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        RequireSuccess(createResult);

        // Add Region as row field
        var addResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        RequireSuccess(addResult);

        // Act
        var descending = _pivotCommands.SortField(batch, "TestPivot", "Region", SortDirection.Descending);
        RequireSuccess(descending);
        RequireSuccess(_pivotCommands.Refresh(batch, "TestPivot"));
        Assert.Equal(["South", "North"], ReadRegions());

        var result = _pivotCommands.SortField(batch, "TestPivot", "Region", SortDirection.Ascending);

        // Assert
        RequireSuccess(result);
        RequireSuccess(_pivotCommands.Refresh(batch, "TestPivot"));
        Assert.Equal(["North", "South"], ReadRegions());
        AssertOriginalSales();

        string[] ReadRegions()
        {
            var data = _pivotCommands.GetData(batch, "TestPivot");
            RequireSuccess(data);
            return data.Values.Select(row => row[0]?.ToString())
                .OfType<string>().Where(value => value is "North" or "South").ToArray();
        }

    }

    private void AssertScalarPivot(double expected)
    {
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, "TestPivot"));
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, "TestPivot")).Values;
        Assert.Equal(expected, Convert.ToDouble(values[^1][^1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void NativePivotCurrency_EscapedLiteral_PreservesDollarDisplay()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));
        ReadNativePivot(_salesSheetName, "TestPivot", pivot =>
        {
            Microsoft.Office.Interop.Excel.PivotFields? fields = null;
            Microsoft.Office.Interop.Excel.PivotField? field = null;
            try
            {
                fields = (Microsoft.Office.Interop.Excel.PivotFields)pivot.DataFields;
                field = fields.Item(1);
                field.NumberFormat = "\\$#,##0.00";
                return 0;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref field);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref fields);
            }
        });
        AssertPivotNumberFormat("$#,##0.00", 650,
            _fixture.ExecuteRawVerification((ctx, _) => $"$650{ctx.FormatTranslator.DecimalSeparator}00"));
    }

    [Theory]
    [InlineData("$#,##0.00", "$#,##0.00")]
    [InlineData("\\$#,##0.00", "$#,##0.00")]
    [InlineData("\"$\"#,##0.00", "$#,##0.00")]
    [InlineData("[$$-409]#,##0.00", "[$$-409]#,##0.00")]
    public void SetFieldFormat_DollarLiteralsAndLocaleMetadata_PreservesDisplay(
        string format, string expectedFormat)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));
        RequireSuccess(_pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", format));
        AssertPivotNumberFormat(expectedFormat, 650,
            _fixture.ExecuteRawVerification((ctx, _) => $"$650{ctx.FormatTranslator.DecimalSeparator}00"));
    }

    [Fact]
    public async Task SetFieldName_ValueCaption_PersistsAfterReopen()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));
        var renamed = RequireSuccess(_pivotCommands.SetFieldName(batch, "TestPivot", "Sales", "Revenue"));
        Assert.Equal(PivotFieldArea.Value, renamed.Area);
        AssertNativeField("TestPivot", "Sales", PivotFieldArea.Value, caption: "Revenue");
        AssertPivotSales(325, 325);

        await _fixture.SaveAndReopenAsync();

        AssertNativeField("TestPivot", "Sales", PivotFieldArea.Value, caption: "Revenue");
        AssertPivotSales(325, 325);
        AssertOriginalSales();
    }

    [Fact]
    public void SetFieldName_RowCaption_UpdatesDisplayedFieldAndPreservesSource()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));
        var renamed = RequireSuccess(_pivotCommands.SetFieldName(batch, "TestPivot", "Region", "Territory"));
        Assert.Equal(PivotFieldArea.Row, renamed.Area);
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Row, caption: "Territory");
        AssertPivotSales(325, 325);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData("rename")]
    [InlineData("function")]
    [InlineData("format")]
    [InlineData("remove")]
    [InlineData("sort")]
    [InlineData("filter")]
    public void FieldMutation_MissingField_PreservesConfiguredPivot(string action)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));
        RequireSuccess(_pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", "$#,##0.00"));
        AssertPivotSales(325, 325);
        var before = SnapshotPivot();
        var dataBefore = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.GetData(batch, "TestPivot")).Values);

        var error = Assert.Throws<InvalidOperationException>(() =>
        {
            const string field = "MissingField";
            ResultBase result = action switch
            {
                "rename" => _pivotCommands.SetFieldName(batch, "TestPivot", field, "Changed"),
                "function" => _pivotCommands.SetFieldFunction(batch, "TestPivot", field, AggregationFunction.Average),
                "format" => _pivotCommands.SetFieldFormat(batch, "TestPivot", field, "0.00%"),
                "remove" => _pivotCommands.RemoveField(batch, "TestPivot", field),
                "sort" => _pivotCommands.SortField(batch, "TestPivot", field, SortDirection.Descending),
                "filter" => _pivotCommands.SetFieldFilter(batch, "TestPivot", field, ["South"]),
                _ => throw new ArgumentOutOfRangeException(nameof(action))
            };
            Assert.Fail($"Expected a thrown rejection, but received Success={result.Success}.");
        });

        Assert.Contains("MissingField", error.Message);
        Assert.Equal(before, SnapshotPivot());
        Assert.Equal(dataBefore, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.GetData(batch, "TestPivot")).Values));
        AssertNativeField("TestPivot", "Region", PivotFieldArea.Row);
        AssertNativeField("TestPivot", "Sales", PivotFieldArea.Value, function: AggregationFunction.Sum);
        AssertOriginalSales();
    }
}
