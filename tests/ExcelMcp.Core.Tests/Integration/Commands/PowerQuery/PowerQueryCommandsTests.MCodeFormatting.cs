using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for M-code formatting feature.
/// Tests verify that M code (Power Query) is automatically formatted on write operations.
/// Write operations (Create, Update) format M code via powerqueryformatter.com API.
/// Read operations (List, View) return M code as stored in the workbook.
/// Note: Formatting may add newlines and spacing, but these tests focus on content preservation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public class PowerQueryCommandsTests_MCodeFormatting : IClassFixture<PowerQueryTestsFixture>
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly PowerQueryTestsFixture _fixture;

    public PowerQueryCommandsTests_MCodeFormatting(PowerQueryTestsFixture fixture)
    {
        _fixture = fixture;
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
    }

    /// <summary>
    /// Tests that Create preserves M code content.
    /// Verifies that the query can be created and the M code is stored correctly.
    /// Formatting may add newlines/spacing, but core content should be preserved.
    /// </summary>
    [Fact]
    public void Create_WithUnformattedMCode_PreservesContent()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_CreateFormatted_{Guid.NewGuid():N}"[..30];

        // Unformatted M code (single line, no spaces, compact)
        var unformattedMCode = "let Source=Excel.CurrentWorkbook(){[Name=\"Table1\"]}[Content],Filtered=Table.SelectRows(Source,each [Column1]>5) in Filtered";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query (should format automatically)
        _powerQueryCommands.Create(batch, queryName, unformattedMCode, PowerQueryLoadMode.ConnectionOnly);

        // Retrieve and verify
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        // The formatted version should contain the core function names
        // (formatting may add whitespace/newlines but preserves content)
        Assert.Contains("let", viewResult.MCode, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Excel.CurrentWorkbook", viewResult.MCode, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Table.SelectRows", viewResult.MCode, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("in", viewResult.MCode, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(viewResult.MCode);
    }

    /// <summary>
    /// Tests that Update preserves M code content.
    /// Verifies that the query can be updated and the M code is stored correctly.
    /// </summary>
    [Fact]
    public void Update_WithUnformattedMCode_PreservesContent()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_UpdateFormatted_{Guid.NewGuid():N}"[..30];

        var originalMCode = @"let
    Source = 1
in
    Source";

        // Unformatted update M code (single line, no spaces)
        var unformattedUpdate = "let Source=#table({\"A\",\"B\"},{{1,2},{3,4}}),Filtered=Table.SelectRows(Source,each [A]>1) in Filtered";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query
        _powerQueryCommands.Create(batch, queryName, originalMCode, PowerQueryLoadMode.ConnectionOnly);

        // Update with unformatted M code (should format automatically)
        _powerQueryCommands.Update(batch, queryName, unformattedUpdate, refresh: false);

        // Retrieve and verify
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        // The formatted version should contain the core function names
        Assert.Contains("#table", viewResult.MCode);
        Assert.Contains("Table.SelectRows", viewResult.MCode);
        Assert.NotEmpty(viewResult.MCode);
    }

    /// <summary>
    /// Tests that pre-formatted M code is preserved/enhanced by the formatter.
    /// Verifies that already-formatted code doesn't break.
    /// </summary>
    [Fact]
    public void Create_WithPreformattedMCode_PreservesReadability()
    {
        var testFile = _fixture.CreateTestFile();
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

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query with pre-formatted M code
        _powerQueryCommands.Create(batch, queryName, preformattedMCode, PowerQueryLoadMode.ConnectionOnly);

        // Retrieve and verify
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        // Should still contain the structure (formatting shouldn't break it)
        Assert.Contains("#table", viewResult.MCode);
        Assert.Contains("ProductID", viewResult.MCode);
        Assert.Contains("let", viewResult.MCode, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("in", viewResult.MCode, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(viewResult.MCode);
    }

    /// <summary>
    /// Tests that empty or whitespace M code is handled gracefully.
    /// The Create operation should fail validation before reaching the formatter.
    /// </summary>
    [Fact]
    public void Create_WithEmptyMCode_ThrowsArgumentException()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_Empty_{Guid.NewGuid():N}"[..30];

        using var batch = ExcelSession.BeginBatch(testFile);

        // Empty M code should fail validation (not reach formatter)
        Assert.Throws<ArgumentException>(() =>
            _powerQueryCommands.Create(batch, queryName, "", PowerQueryLoadMode.ConnectionOnly));

        // Whitespace-only M code should also fail
        Assert.Throws<ArgumentException>(() =>
            _powerQueryCommands.Create(batch, queryName, "   ", PowerQueryLoadMode.ConnectionOnly));
    }

    /// <summary>
    /// Tests that View returns M code as stored (formatter applied on write).
    /// Verifies that read operations don't re-format.
    /// </summary>
    [Fact]
    public void View_AfterCreate_ReturnsMCodeAsStored()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_ViewStored_{Guid.NewGuid():N}"[..30];

        // Simple M code
        var mCode = "let x=1 in x";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // View twice - should return same result each time
        var viewResult1 = _powerQueryCommands.View(batch, queryName);
        var viewResult2 = _powerQueryCommands.View(batch, queryName);

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
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_ListQuery_{Guid.NewGuid():N}"[..30];

        // Unformatted M code
        var unformattedMCode = "let Source=1,Result=Source+1 in Result";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query
        _powerQueryCommands.Create(batch, queryName, unformattedMCode, PowerQueryLoadMode.ConnectionOnly);

        // List queries
        var listResult = _powerQueryCommands.List(batch);
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
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_Complex_{Guid.NewGuid():N}"[..30];

        // Complex unformatted M code with multiple steps
        var complexMCode = "let Source=#table({\"A\",\"B\",\"C\"},{{1,2,3},{4,5,6},{7,8,9}}),Filtered=Table.SelectRows(Source,each [A]>3),Transformed=Table.TransformColumnTypes(Filtered,{{\"A\",type number},{\"B\",type number},{\"C\",type number}}),Added=Table.AddColumn(Transformed,\"Sum\",each [A]+[B]+[C]) in Added";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query (should format automatically)
        _powerQueryCommands.Create(batch, queryName, complexMCode, PowerQueryLoadMode.ConnectionOnly);

        // Retrieve and verify
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        // Verify all steps are present
        Assert.Contains("#table", viewResult.MCode);
        Assert.Contains("Table.SelectRows", viewResult.MCode);
        Assert.Contains("Table.TransformColumnTypes", viewResult.MCode);
        Assert.Contains("Table.AddColumn", viewResult.MCode);
        Assert.NotEmpty(viewResult.MCode);

        // Verify the query can still be viewed without errors (formatting didn't corrupt it)
        var verifyResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(verifyResult.Success);
    }

    /// <summary>
    /// Tests that sequential Create and Update operations both preserve content.
    /// Verifies that formatting is consistent across operations.
    /// </summary>
    [Fact]
    public void CreateThenUpdate_BothOperationsPreserveContent()
    {
        var testFile = _fixture.CreateTestFile();
        var queryName = $"Test_Sequential_{Guid.NewGuid():N}"[..30];

        // First unformatted M code
        var createMCode = "let x=1,y=2 in x+y";

        // Second unformatted M code
        var updateMCode = "let a=10,b=20,c=30 in a+b+c";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query
        _powerQueryCommands.Create(batch, queryName, createMCode, PowerQueryLoadMode.ConnectionOnly);

        var afterCreate = _powerQueryCommands.View(batch, queryName);
        Assert.True(afterCreate.Success);
        Assert.NotEmpty(afterCreate.MCode);

        // Update query
        _powerQueryCommands.Update(batch, queryName, updateMCode, refresh: false);

        var afterUpdate = _powerQueryCommands.View(batch, queryName);
        Assert.True(afterUpdate.Success);
        Assert.NotEmpty(afterUpdate.MCode);

        // New content should be present
        Assert.Contains("a", afterUpdate.MCode);
        Assert.Contains("b", afterUpdate.MCode);
        Assert.Contains("c", afterUpdate.MCode);

        // Old unique values should be replaced (x=1, y=2)
        Assert.DoesNotContain("x = 1", afterUpdate.MCode);
        Assert.DoesNotContain("y = 2", afterUpdate.MCode);
    }
}
