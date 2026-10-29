using System.Globalization;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for Parameter value operations (get, set)
/// </summary>
public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void Create_InvalidFormula_DoesNotLeaveTemporaryName()
    {
        var batch = _fixture.BatchToken;
        var name = CreateUniqueNamedRangeName();
        var before = _parameterCommands.List(batch);
        Assert.True(before.Success, before.ErrorMessage);

        Assert.Throws<InvalidOperationException>(() => _parameterCommands.Create(batch, name, "=("));

        var after = _parameterCommands.List(batch);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.NamedRanges.Select(item => item.Name), after.NamedRanges.Select(item => item.Name));
        Assert.DoesNotContain(after.NamedRanges, item =>
            item.Name.StartsWith("ExcelMcpTemp_", StringComparison.Ordinal));
    }

    [Fact]
    public void CreateAndUpdate_FormulaReferences_UseInvariantSyntaxUnderGermanCulture()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        var originalUiCulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("de-DE");

            var batch = _fixture.BatchToken;
            var sheetName = _fixture.CreateNamedTestSheet(batch, $"Formula names {Guid.NewGuid():N}"[..27]);
            var tableName = $"FormulaTable_{Guid.NewGuid():N}";
            var rateName = CreateUniqueNamedRangeName();
            var totalName = CreateUniqueNamedRangeName();
            Assert.True(_commands.SetValues(batch, sheetName, "A1:B3",
                [["Item", "Total"], ["A", 3], ["B", 6.75]]).Success);
            var table = _fixture.Send("table.create", new
            {
                sheetName,
                tableName,
                rangeAddress = "A1:B3",
                hasHeaders = true
            });
            Assert.True(table.Success, table.ErrorMessage);
            _fixture.RegisterTableForCleanup(tableName);

            Assert.True(_parameterCommands.Create(batch, rateName, "=0.19").Success);
            _fixture.RegisterNamedRangeForCleanup(rateName);
            Assert.True(_commands.SetFormulas(batch, sheetName, "D1", [[$"={rateName}"]]).Success);
            var rateValue = _commands.GetValues(batch, sheetName, "D1");
            Assert.True(rateValue.Success, rateValue.ErrorMessage);
            Assert.Equal(0.19, Convert.ToDouble(rateValue.Values[0][0], CultureInfo.InvariantCulture));

            Assert.True(_parameterCommands.Create(batch, totalName, $"=SUM({tableName}[Total])").Success);
            _fixture.RegisterNamedRangeForCleanup(totalName);
            Assert.True(_commands.SetFormulas(batch, sheetName, "D2", [[$"={totalName}"]]).Success);
            var totalValue = _commands.GetValues(batch, sheetName, "D2");
            Assert.True(totalValue.Success, totalValue.ErrorMessage);
            Assert.Equal(9.75, Convert.ToDouble(totalValue.Values[0][0], CultureInfo.InvariantCulture));

            Assert.True(_parameterCommands.Update(batch, rateName, "=1.25").Success);
            var updatedValue = _commands.GetValues(batch, sheetName, "D1");
            Assert.True(updatedValue.Success, updatedValue.ErrorMessage);
            Assert.Equal(1.25, Convert.ToDouble(updatedValue.Values[0][0], CultureInfo.InvariantCulture));

            Assert.True(_parameterCommands.Update(batch, rateName, $"=\"[\"&'{sheetName}'!$B$2").Success);
            var textValue = _commands.GetValues(batch, sheetName, "D1");
            Assert.True(textValue.Success, textValue.ErrorMessage);
            Assert.Equal("[3", textValue.Values[0][0]);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public void ReadAndList_MatrixNamedRange_PreserveValuesAndDimensions()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateNamedTestSheet(batch, $"Matrix values {Guid.NewGuid():N}"[..27]);
        var name = CreateUniqueNamedRangeName();
        var reference = $"'{sheetName}'!$B$2:$D$3";
        Assert.True(_commands.SetValues(batch, sheetName, "B2:D3",
            [["Header", 12.5, true], ["Second", null, -7d]]).Success);
        Assert.True(_parameterCommands.Create(batch, name, reference).Success);
        _fixture.RegisterNamedRangeForCleanup(name);

        var read = _parameterCommands.Read(batch, name.ToLowerInvariant());
        Assert.Equal(name.ToLowerInvariant(), read.Name);
        Assert.Equal($"={reference}", read.RefersTo);
        Assert.Equal("Array", read.ValueType);
        AssertMatrix(read.Value);

        var listed = _parameterCommands.List(batch);
        Assert.True(listed.Success, listed.ErrorMessage);
        var entry = Assert.Single(listed.NamedRanges, item => item.Name == name);
        Assert.Equal($"={reference}", entry.RefersTo);
        Assert.Equal("Array", entry.ValueType);
        Assert.Equal(6, entry.CellCount);
        Assert.Null(entry.ValueOmittedReason);
        AssertMatrix(entry.Value);

        static void AssertMatrix(object? value)
        {
            var rows = Assert.IsType<JsonElement>(value);
            Assert.Equal(JsonValueKind.Array, rows.ValueKind);
            Assert.Equal(2, rows.GetArrayLength());
            Assert.Equal(3, rows[0].GetArrayLength());
            Assert.Equal(3, rows[1].GetArrayLength());
            Assert.Equal("Header", rows[0][0].GetString());
            Assert.Equal(12.5, rows[0][1].GetDouble());
            Assert.True(rows[0][2].GetBoolean());
            Assert.Equal("Second", rows[1][0].GetString());
            Assert.Equal(JsonValueKind.Null, rows[1][1].ValueKind);
            Assert.Equal(-7d, rows[1][2].GetDouble());
        }
    }

    /// <inheritdoc/>
    [Fact]
    public void Set_ExistingParameter_UpdatesValue()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create parameter first
        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [["Original", "Untouched"]]).Success);
        Assert.Equal("Original", _parameterCommands.Read(batch, paramName).Value);

        // Set the parameter value
        Assert.True(_parameterCommands.Write(batch, paramName, "TestValue").Success);

        // Assert - Verify the parameter value was actually set by reading it back
        var namedRangeValue = _parameterCommands.Read(batch, paramName);
        Assert.Equal("TestValue", namedRangeValue.Value);
        Assert.Equal(paramName, namedRangeValue.Name);
        Assert.Equal("String", namedRangeValue.ValueType);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["TestValue", "Untouched"], Assert.Single(cells.Values));
    }

    [Fact]
    public void Write_DottedIdentifier_PreservesText()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_parameterCommands.Write(batch, paramName, "2.0.13").Success);

        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        Assert.Equal("2.0.13", namedRangeValue.Value);
        Assert.Equal("String", namedRangeValue.ValueType);
        Assert.Equal(paramName, namedRangeValue.Name);
        var cells = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal("2.0.13", Assert.Single(Assert.Single(cells.Values)));
    }

    [Theory]
    [InlineData("123.45", 123.45, "Double")]
    [InlineData("true", true, "Boolean")]
    public void Write_InvariantScalar_StoresTypedValue(
        string input,
        object expectedValue,
        string expectedType)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_parameterCommands.Write(batch, paramName, input).Success);

        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        Assert.Equal(expectedValue, namedRangeValue.Value);
        Assert.Equal(expectedType, namedRangeValue.ValueType);
        Assert.Equal(paramName, namedRangeValue.Name);
        var cells = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(expectedValue, Assert.Single(Assert.Single(cells.Values)));
    }

    /// <inheritdoc/>

    [Fact]
    public void Get_ExistingParameter_ReturnsValue()
    {
        // Arrange
        string testValue = "Integration Test Value";
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create and set parameter value
        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[testValue]]).Success);

        // Get the parameter value
        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        // Assert
        Assert.Equal(testValue, namedRangeValue.Value);
        Assert.Equal(paramName, namedRangeValue.Name);
        Assert.Equal("String", namedRangeValue.ValueType);
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_WithNonExistentParameter_ThrowsException()
    {
        // Arrange & Act & Assert
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var exception = Assert.Throws<InvalidOperationException>(
            () => _parameterCommands.Read(batch, $"NonExistent_{Guid.NewGuid():N}"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
