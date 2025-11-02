using Xunit;
using Xunit.Abstractions;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// CRITICAL REGRESSION TESTS: Verify Success flag is NEVER true when ErrorMessage is set
/// Bug: Success=true with ErrorMessage confuses LLMs and causes silent failures
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("BugFix", "SuccessWithError")]
public class PowerQuerySuccessErrorRegressionTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IPowerQueryCommands _commands;
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;

    public PowerQuerySuccessErrorRegressionTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        var dataModelCommands = new DataModelCommands();
        _commands = new PowerQueryCommands(dataModelCommands);
        _tempDir = fixture.TempDir;
        _output = output;
    }

    /// <summary>
    /// CRITICAL: If import fails to load to data model, Success MUST be false
    /// </summary>
    [Fact]
    public async Task Import_LoadToDataModelFails_SuccessIsFalse()
    {
        // Arrange - Create query that references non-existent table
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQuerySuccessErrorRegressionTests),
            nameof(Import_LoadToDataModelFails_SuccessIsFalse),
            _tempDir);

        var queryFile = Path.Combine(_tempDir, "bad_query.pq");
        var mCode = @"
let
    Source = Excel.CurrentWorkbook(){[Name=""NonExistentTable""]}[Content]
in
    Source
";
        await System.IO.File.WriteAllTextAsync(queryFile, mCode);

        // Act - Try to import with data-model destination
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.ImportAsync(batch, "BadQuery", queryFile, "data-model");

        // Assert - CRITICAL: Success MUST be false when ErrorMessage is set
        _output.WriteLine($"Success: {result.Success}");
        _output.WriteLine($"ErrorMessage: {result.ErrorMessage}");

        Assert.False(result.Success, 
            "CRITICAL BUG: Success=true but ErrorMessage is set! This confuses LLMs.");
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage),
            "ErrorMessage should be set when load fails");
        Assert.Contains("failed to load", result.ErrorMessage.ToLowerInvariant());
    }

    /// <summary>
    /// SUCCESS should be true ONLY when operation completes without errors
    /// </summary>
    [Fact]
    public async Task Import_Success_NoErrorMessage()
    {
        // Arrange - Create valid query
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQuerySuccessErrorRegressionTests),
            nameof(Import_Success_NoErrorMessage),
            _tempDir);

        var queryFile = Path.Combine(_tempDir, "valid_query.pq");
        var mCode = @"
let
    Source = {1, 2, 3}
in
    Source
";
        await System.IO.File.WriteAllTextAsync(queryFile, mCode);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.ImportAsync(batch, "ValidQuery", queryFile, "worksheet");

        // Assert - Success with NO error message
        Assert.True(result.Success, "Valid import should succeed");
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage),
            "Successful operation should NOT have ErrorMessage");
    }

    /// <summary>
    /// Verify the invariant: Success=true implies ErrorMessage is null/empty
    /// </summary>
    [Theory]
    [InlineData("worksheet")]
    [InlineData("connection-only")]
    public async Task Import_SuccessImpliesNoError_AllDestinations(string destination)
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQuerySuccessErrorRegressionTests),
            $"{nameof(Import_SuccessImpliesNoError_AllDestinations)}_{destination}",
            _tempDir);

        var queryFile = Path.Combine(_tempDir, $"query_{destination}.pq");
        var mCode = "let Source = {1, 2, 3} in Source";
        await System.IO.File.WriteAllTextAsync(queryFile, mCode);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.ImportAsync(batch, $"Query_{destination}", queryFile, destination);

        // Assert - Invariant: Success=true => ErrorMessage is null/empty
        if (result.Success)
        {
            Assert.True(string.IsNullOrEmpty(result.ErrorMessage),
                $"INVARIANT VIOLATION: Success=true but ErrorMessage='{result.ErrorMessage}' for destination '{destination}'");
        }
    }
}
