using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryFormattingTests
{
    [Fact]
    public void Create_WithUnformattedMCode_PreservesExactInputByDefault()
    {
        var queryName = $"Test_CreateFormatted_{Guid.NewGuid():N}"[..30];

        // Unformatted M code (single line, no spaces, compact)
        var unformattedMCode = "let Source=Excel.CurrentWorkbook(){[Name=\"Table1\"]}[Content],Filtered=Table.SelectRows(Source,each [Column1]>5) in Filtered";

        var batch = _fixture.BatchToken;

        // Create query without remote formatting opt-in
        RequireSuccess(_queries.Create(batch, queryName, unformattedMCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Retrieve and verify
        AssertStoredM(queryName, unformattedMCode);
    }

    /// <summary>
    /// Tests that Update preserves M code exactly by default.
    /// </summary>
    [Fact]
    public void Update_WithUnformattedMCode_PreservesExactInputByDefault()
    {
        var queryName = $"Test_UpdateFormatted_{Guid.NewGuid():N}"[..30];

        var originalMCode = @"let
    Source = 1
in
    Source";

        // Unformatted update M code (single line, no spaces)
        var unformattedUpdate = "let Source=#table({\"A\",\"B\"},{{1,2},{3,4}}),Filtered=Table.SelectRows(Source,each [A]>1) in Filtered";

        var batch = _fixture.BatchToken;

        // Create query
        RequireSuccess(_queries.Create(batch, queryName, originalMCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        AssertStoredM(queryName, originalMCode);

        // Update without remote formatting opt-in
        RequireSuccess(_queries.Update(batch, queryName, unformattedUpdate, refresh: false));

        // Retrieve and verify
        AssertStoredM(queryName, unformattedUpdate);
        AssertEvaluatedM(unformattedUpdate, ["A", "B"], [[3, 4]]);
        AssertStoredM(queryName, unformattedUpdate);
    }

    /// <summary>
    /// Tests that pre-formatted M code is preserved exactly by default.
    /// </summary>
    [Fact]
    public void Create_WithPreformattedMCode_PreservesReadability()
    {
        var queryName = $"Test_PreFormatted_{Guid.NewGuid():N}"[..30];

        // Pre-formatted M code (with proper indentation)
        var preformattedMCode = @"let
    Source = #table(
        {""ProductID"", ""ProductName"", ""Price""},
        {
            {1, ""Widget"", 10.99},
            {2, ""Gadget"", 25.50}
        }
    )
in
    Source";

        var batch = _fixture.BatchToken;

        // Create query with pre-formatted M code
        RequireSuccess(_queries.Create(batch, queryName, preformattedMCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Retrieve and verify
        AssertStoredM(queryName, preformattedMCode);
        AssertEvaluatedM(preformattedMCode, ["ProductID", "ProductName", "Price"],
            [[1, "Widget", 10.99], [2, "Gadget", 25.50]]);
        AssertStoredM(queryName, preformattedMCode);
    }

    /// <summary>
    /// Tests that empty or whitespace M code is handled gracefully.
    /// The Create operation should fail validation before reaching the formatter.
    /// </summary>
    [Fact]
    public void Create_WithEmptyMCode_ThrowsArgumentException()
    {
        var queryName = $"Test_Empty_{Guid.NewGuid():N}"[..30];

        var batch = _fixture.BatchToken;
        var guardName = $"FormatGuard_{Guid.NewGuid():N}"[..30];
        const string guardM = "#table(type table [A = number, B = number], {{7,14},{9,18}})";
        RequireSuccess(_queries.Create(batch, guardName, guardM,
            PowerQueryLoadMode.LoadToTable, guardName));
        _fixture.RegisterPowerQueryForCleanup(guardName);
        _fixture.RegisterSheetForCleanup(guardName);
        Assert.Equal(guardM, RequireSuccess(_queries.View(batch, guardName)).MCode);
        var expectedCells = new List<List<object?>> { new() { "A", "B" }, new() { 7, 14 }, new() { 9, 18 } };
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(expectedCells),
            System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(_commands.GetValues(batch, guardName, "A1:B3")).Values));
        var before = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_queries.List(batch)).Queries);
        var loadBefore = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_queries.GetLoadConfig(batch, guardName)));
        foreach (var invalidM in new[] { "", "   " })
        {
            var error = Assert.Throws<ArgumentException>(() =>
                _queries.Create(batch, queryName, invalidM, PowerQueryLoadMode.ConnectionOnly));
            Assert.Contains("mCode", error.Message, StringComparison.Ordinal);
            Assert.Contains("InvalidInput/ArgumentException", error.Message, StringComparison.Ordinal);
            Assert.Equal(before, System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(_queries.List(batch)).Queries));
            Assert.Equal(guardM, RequireSuccess(_queries.View(batch, guardName)).MCode);
            Assert.Equal(loadBefore, System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(_queries.GetLoadConfig(batch, guardName))));
            Assert.Equal(System.Text.Json.JsonSerializer.Serialize(expectedCells),
                System.Text.Json.JsonSerializer.Serialize(
                    RequireSuccess(_commands.GetValues(batch, guardName, "A1:B3")).Values));
        }
    }

    /// <summary>
    /// Tests that View returns M code as stored.
    /// Verifies that read operations don't re-format.
    /// </summary>
    [Fact]
    public void View_AfterCreate_ReturnsMCodeAsStored()
    {
        var queryName = $"Test_ViewStored_{Guid.NewGuid():N}"[..30];

        // Simple M code
        var mCode = "let x=1 in x";

        var batch = _fixture.BatchToken;

        // Create query
        RequireSuccess(_queries.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // View twice - should return same result each time
        AssertStoredM(queryName, mCode);
        AssertStoredM(queryName, mCode);
    }

    /// <summary>
    /// Tests that complex M code with multiple steps is properly handled.
    /// Verifies that multi-step queries maintain their structure.
    /// </summary>
    [Fact]
    public void Create_WithComplexMultiStepMCode_PreservesAllSteps()
    {
        var queryName = $"Test_Complex_{Guid.NewGuid():N}"[..30];

        // Complex unformatted M code with multiple steps
        var complexMCode = "let Source=#table({\"A\",\"B\",\"C\"},{{1,2,3},{4,5,6},{7,8,9}}),Filtered=Table.SelectRows(Source,each [A]>3),Transformed=Table.TransformColumnTypes(Filtered,{{\"A\",type number},{\"B\",type number},{\"C\",type number}}),Added=Table.AddColumn(Transformed,\"Sum\",each [A]+[B]+[C]) in Added";

        var batch = _fixture.BatchToken;

        // Create query without remote formatting opt-in
        RequireSuccess(_queries.Create(batch, queryName, complexMCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Retrieve and verify
        AssertStoredM(queryName, complexMCode);
        AssertEvaluatedM(complexMCode, ["A", "B", "C", "Sum"], [[4, 5, 6, 15], [7, 8, 9, 24]]);
        AssertStoredM(queryName, complexMCode);
    }

    /// <summary>
    /// Tests that sequential Create and Update operations both preserve content.
    /// Verifies that formatting is consistent across operations.
    /// </summary>
    [Fact]
    public void CreateThenUpdate_BothOperationsPreserveContent()
    {
        var queryName = $"Test_Sequential_{Guid.NewGuid():N}"[..30];

        // First unformatted M code
        var createMCode = "let x=1,y=2 in x+y";

        // Second unformatted M code
        var updateMCode = "let a=10,b=20,c=30 in a+b+c";

        var batch = _fixture.BatchToken;

        // Create query
        RequireSuccess(_queries.Create(batch, queryName, createMCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);

        AssertStoredM(queryName, createMCode);

        // Update query
        RequireSuccess(_queries.Update(batch, queryName, updateMCode, refresh: false));

        AssertStoredM(queryName, updateMCode);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Update_WithEmptyMCode_PreservesStoredQuery(string invalidM)
    {
        var name = $"FormatReject_{Guid.NewGuid():N}"[..30];
        const string mCode = "#table({\"Value\"}, {{17},{29}})";
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, mCode, PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(name);
        AssertStoredM(name, mCode);
        var before = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_queries.List(_fixture.BatchToken)).Queries);
        var error = Assert.Throws<ArgumentException>(() =>
            _queries.Update(_fixture.BatchToken, name, invalidM, refresh: false));
        Assert.Contains("mCode", error.Message, StringComparison.Ordinal);
        Assert.Contains("InvalidInput/ArgumentException", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_queries.List(_fixture.BatchToken)).Queries));
        AssertStoredM(name, mCode);
        AssertEvaluatedM(mCode, ["Value"], [[17], [29]]);
        AssertStoredM(name, mCode);
    }

    private void AssertStoredM(string name, string mCode)
    {
        var view = RequireSuccess(_queries.View(_fixture.BatchToken, name));
        Assert.Equal(name, view.QueryName);
        Assert.Equal(mCode, view.MCode);
        Assert.Equal(mCode.Length, view.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, view.LoadMode);
        Assert.True(view.IsConnectionOnly);
        Assert.False(view.HasConnection);
        Assert.False(view.IsLoadedToDataModel);
        Assert.Null(view.TargetSheet);
        var listed = Assert.Single(RequireSuccess(_queries.List(_fixture.BatchToken)).Queries,
            query => query.Name == name);
        Assert.Equal(mCode.Length > 80 ? mCode[..77] + "..." : mCode, listed.FormulaPreview);
        Assert.Equal(mCode.Length, listed.CharacterCount);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, listed.LoadMode);
        Assert.True(listed.IsConnectionOnly);
        Assert.False(listed.IsLoadedToDataModel);
        Assert.Null(listed.TargetSheet);
    }

    private void AssertEvaluatedM(string mCode, string[] columns, object[][] rows)
    {
        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, mCode));
        Assert.Equal(mCode, result.MCode);
        Assert.Equal(columns, result.Columns);
        Assert.Equal(columns.Length, result.ColumnCount);
        Assert.Equal(rows.Length, result.RowCount);
        Assert.Equal(rows.Length, result.Rows.Count);
        for (var row = 0; row < rows.Length; row++)
        {
            Assert.Equal(columns.Length, result.Rows[row].Count);
            for (var column = 0; column < columns.Length; column++)
            {
                if (rows[row][column] is string text)
                {
                    Assert.Equal(text, Assert.IsType<string>(result.Rows[row][column]));
                }
                else
                {
                    Assert.Equal(Convert.ToDecimal(rows[row][column], System.Globalization.CultureInfo.InvariantCulture),
                        Convert.ToDecimal(result.Rows[row][column], System.Globalization.CultureInfo.InvariantCulture));
                }
            }
        }
    }
}
