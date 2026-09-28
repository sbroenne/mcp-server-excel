using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
public sealed class FixtureProcessOwnershipTests(ITestOutputHelper output)
{
    [Fact]
    public void OwnedScope_ReportsOwnedLeak_AndPreservesIndependentExcel()
    {
        var directory = Path.Join(Path.GetTempPath(), $"ScopeOwnership_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            using var independent = ExcelBatch.CreateNewWorkbook(Path.Join(directory, "independent.xlsx"), isMacroEnabled: false);
            using var ownership = new OwnedExcelProcessScope();
            using var owned = ExcelBatch.CreateNewWorkbook(Path.Join(directory, "owned.xlsx"), isMacroEnabled: false);

            var error = Assert.Throws<Xunit.Sdk.TrueException>(() => ownership.AssertAllExited());

            Assert.Contains("survived shutdown", error.Message, StringComparison.Ordinal);
            Assert.False(owned.IsExcelProcessAlive());
            Assert.True(independent.IsExcelProcessAlive());
            independent.Execute((context, _) => Assert.Equal("independent.xlsx", context.Book.Name));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("session-manager")]
    [InlineData("excel-session")]
    [InlineData("disposal")]
    public async Task FixtureSetup_PreservesIndependentExcel(string fixture)
    {
        var path = Path.Join(Path.GetTempPath(), $"FixtureOwnership_{Guid.NewGuid():N}.xlsx");
        try
        {
            using var independent = ExcelBatch.CreateNewWorkbook(path, isMacroEnabled: false);
            switch (fixture)
            {
                case "session-manager":
                    using (var instance = new SessionManagerTests(output)) { }
                    break;
                case "excel-session":
                    using (var instance = new ExcelSessionTests(output)) { }
                    break;
                case "disposal":
                    var disposal = new DisposalVerificationTest(output);
                    try
                    {
                        await disposal.InitializeAsync();
                    }
                    finally
                    {
                        await disposal.DisposeAsync();
                    }
                    break;
            }

            Assert.True(independent.IsExcelProcessAlive(), $"{fixture} must preserve independently owned Excel.");
            independent.Execute((context, _) => Assert.Equal(Path.GetFileName(path), context.Book.Name));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
