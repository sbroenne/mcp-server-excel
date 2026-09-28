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
        _queries.Create(batch, queryName, unformattedMCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Retrieve and verify
        var viewResult = _queries.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        Assert.Equal(unformattedMCode, viewResult.MCode);
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
        _queries.Create(batch, queryName, originalMCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Update without remote formatting opt-in
        _queries.Update(batch, queryName, unformattedUpdate, refresh: false);

        // Retrieve and verify
        var viewResult = _queries.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        Assert.Equal(unformattedUpdate, viewResult.MCode);
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
        _queries.Create(batch, queryName, preformattedMCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Retrieve and verify
        var viewResult = _queries.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        Assert.Equal(preformattedMCode, viewResult.MCode);
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

        // Empty M code should fail validation (not reach formatter)
        Assert.Throws<ArgumentException>(() =>
            _queries.Create(batch, queryName, "", PowerQueryLoadMode.ConnectionOnly));

        // Whitespace-only M code should also fail
        Assert.Throws<ArgumentException>(() =>
            _queries.Create(batch, queryName, "   ", PowerQueryLoadMode.ConnectionOnly));
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
        _queries.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // View twice - should return same result each time
        var viewResult1 = _queries.View(batch, queryName);
        var viewResult2 = _queries.View(batch, queryName);

        Assert.True(viewResult1.Success);
        Assert.True(viewResult2.Success);

        // M code should be identical on both reads (no re-formatting on read)
        Assert.Equal(viewResult1.MCode, viewResult2.MCode);
    }

    /// <summary>
    /// Tests that List returns queries with M code intact.
    /// Verifies that list operation doesn't affect stored M code.
    /// </summary>
    [Fact]
    public void List_AfterCreate_ReturnsQueryWithMCode()
    {
        var queryName = $"Test_ListQuery_{Guid.NewGuid():N}"[..30];

        // Unformatted M code
        var unformattedMCode = "let Source=1,Result=Source+1 in Result";

        var batch = _fixture.BatchToken;

        // Create query
        _queries.Create(batch, queryName, unformattedMCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // List queries
        var listResult = _queries.List(batch);
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");

        // Find our query
        var query = listResult.Queries.FirstOrDefault(q => q.Name == queryName);
        Assert.NotNull(query);

        // Verify query has M code preview
        Assert.False(string.IsNullOrEmpty(query.FormulaPreview));
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
        _queries.Create(batch, queryName, complexMCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        // Retrieve and verify
        var viewResult = _queries.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        Assert.Equal(complexMCode, viewResult.MCode);

        // Verify the query can still be viewed without errors (formatting didn't corrupt it)
        var verifyResult = _queries.View(batch, queryName);
        Assert.True(verifyResult.Success);
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
        _queries.Create(batch, queryName, createMCode, PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var afterCreate = _queries.View(batch, queryName);
        Assert.True(afterCreate.Success);
        Assert.Equal(createMCode, afterCreate.MCode);

        // Update query
        _queries.Update(batch, queryName, updateMCode, refresh: false);

        var afterUpdate = _queries.View(batch, queryName);
        Assert.True(afterUpdate.Success);
        Assert.Equal(updateMCode, afterUpdate.MCode);
    }
}
