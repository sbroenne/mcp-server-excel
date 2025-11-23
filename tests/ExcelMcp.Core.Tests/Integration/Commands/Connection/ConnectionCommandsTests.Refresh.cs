using System.IO;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void Refresh_ConnectionNotFound_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_ConnectionNotFound_ThrowsException),
            _tempDir);

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<InvalidOperationException>(() => _commands.Refresh(batch, "NonExistentConnection"));
        Assert.Contains("not found", exception.Message);
    }

    /// <summary>
    /// Tests refreshing an ACE OLEDB connection bound to an Excel workbook data source.
    /// </summary>
    [Fact]
    public void Refresh_AceOleDbConnection_ReturnsSuccess()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection(
            nameof(Refresh_AceOleDbConnection_ReturnsSuccess));

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            var loadResult = _commands.LoadTo(batch, connectionName, "ProductsData");
            Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");

            var refreshResult = _commands.Refresh(batch, connectionName);
            Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        }
        finally
        {
            if (System.IO.File.Exists(sourceWorkbook))
            {
                System.IO.File.Delete(sourceWorkbook);
            }
        }
    }

    /// <summary>
    /// Tests refreshing an ACE OLEDB connection after modifying the external workbook.
    /// </summary>
    [Fact]
    public void Refresh_AceOleDbConnectionAfterDataUpdate_ReturnsSuccess()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection(
            nameof(Refresh_AceOleDbConnectionAfterDataUpdate_ReturnsSuccess));

        try
        {
            using (var batch = ExcelSession.BeginBatch(testFile))
            {
                var loadResult = _commands.LoadTo(batch, connectionName, "ProductsData");
                Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");
                batch.Save();
            }

            AceOleDbTestHelper.UpdateExcelDataSource(sourceWorkbook, sheet =>
            {
                sheet.Range["B2"].Value2 = 49.99;
            });

            using var refreshBatch = ExcelSession.BeginBatch(testFile);
            var refreshResult = _commands.Refresh(refreshBatch, connectionName);
            Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        }
        finally
        {
            if (System.IO.File.Exists(sourceWorkbook))
            {
                System.IO.File.Delete(sourceWorkbook);
            }
        }
    }

    private (string testFile, string sourceWorkbook, string connectionName) SetupAceOleDbConnection(string testName)
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            testName,
            _tempDir);

        var sourceWorkbook = Path.Combine(_tempDir, $"{testName}_Source.xlsx");
        AceOleDbTestHelper.CreateExcelDataSource(sourceWorkbook);

        var connectionName = "TestAceOleDbConnection";
        ConnectionTestHelper.CreateAceOleDbConnection(testFile, connectionName, sourceWorkbook);

        return (testFile, sourceWorkbook, connectionName);
    }
}
