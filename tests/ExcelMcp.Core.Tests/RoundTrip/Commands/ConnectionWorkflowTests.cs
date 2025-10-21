using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.RoundTrip.Commands;

/// <summary>
/// Round trip tests for connection management workflows.
/// These tests verify complete end-to-end scenarios with real Excel files.
/// </summary>
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ConnectionWorkflowTests : IDisposable
{
    private readonly string _testDirectory;
    private readonly List<string> _testFiles;
    private readonly ConnectionCommands _commands;

    public ConnectionWorkflowTests()
    {
        _testDirectory = Path.Combine(Path.GetTempPath(), $"ExcelMcpConnectionRoundTrip_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDirectory);
        _testFiles = new List<string>();
        _commands = new ConnectionCommands();
    }

    public void Dispose()
    {
        // Clean up test files
        foreach (var file in _testFiles.Where(File.Exists))
        {
            try { File.Delete(file); } catch { /* Ignore cleanup errors */ }
        }
        
        if (Directory.Exists(_testDirectory))
        {
            try { Directory.Delete(_testDirectory, recursive: true); } catch { /* Ignore cleanup errors */ }
        }
        
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void CompleteWorkflow_ExportModifyUpdateVerify_Success()
    {
        // Arrange: Create workbook with connection
        var workbookPath = CreateTestWorkbookWithConnection();
        var exportPath = Path.Combine(_testDirectory, "connection-export.json");
        var modifiedPath = Path.Combine(_testDirectory, "connection-modified.json");
        var connectionName = "TestConnection";

        try
        {
            // Act 1: Export connection to JSON
            var exportResult = _commands.Export(workbookPath, connectionName, exportPath);
            Assert.True(exportResult.Success, $"Export failed: {exportResult.ErrorMessage}");
            Assert.True(File.Exists(exportPath), "Export file not created");

            // Act 2: Modify JSON definition (simulate editing)
            var jsonContent = File.ReadAllText(exportPath);
            Assert.Contains("TestConnection", jsonContent);
            
            // Create a modified version with different description
            // Replace the existing description (from ConnectionTestHelper) with our modified one
            var modifiedJson = jsonContent.Replace(
                "Test web connection created by CreateWebConnection", 
                "Modified via round trip test");
            File.WriteAllText(modifiedPath, modifiedJson);

            // Act 3: Update connection from modified JSON
            var updateResult = _commands.Update(workbookPath, connectionName, modifiedPath);
            Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

            // Act 4: View connection to verify changes
            var viewResult = _commands.View(workbookPath, connectionName);
            Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
            Assert.NotNull(viewResult.DefinitionJson);
            Assert.Contains("Modified via round trip test", viewResult.DefinitionJson);

            // Act 5: Test connection (validate it still works)
            var testResult = _commands.Test(workbookPath, connectionName);
            Assert.True(testResult.Success, $"Test failed: {testResult.ErrorMessage}");

            // Act 6: Get properties to verify configuration preserved
            var propsResult = _commands.GetProperties(workbookPath, connectionName);
            Assert.True(propsResult.Success, $"GetProperties failed: {propsResult.ErrorMessage}");
            // BackgroundQuery and RefreshOnFileOpen are bool, not bool?

            // Act 7: Delete connection (cleanup)
            var deleteResult = _commands.Delete(workbookPath, connectionName);
            Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

            // Verify: Connection no longer exists
            var listResult = _commands.List(workbookPath);
            Assert.True(listResult.Success);
            Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
        }
        finally
        {
            // Cleanup
            if (File.Exists(exportPath)) File.Delete(exportPath);
            if (File.Exists(modifiedPath)) File.Delete(modifiedPath);
        }
    }

    [Fact]
    public void ConnectionRefresh_LoadToWorksheetRefreshVerifyData_Success()
    {
        // Arrange: Create workbook with text file connection
        var workbookPath = CreateTestWorkbookWithTextConnection();
        var sheetName = "DataSheet";
        var connectionName = "DataConnection";

        // Act 1: List connections to verify it exists
        var listResult = _commands.List(workbookPath);
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);

        // Act 2: Load connection to worksheet
        var loadResult = _commands.LoadTo(workbookPath, connectionName, sheetName);
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");

        // Act 3: Get connection properties
        var propsResult = _commands.GetProperties(workbookPath, connectionName);
        Assert.True(propsResult.Success, $"GetProperties failed: {propsResult.ErrorMessage}");
        var originalBackgroundQuery = propsResult.BackgroundQuery;

        // Act 4: Modify connection properties
        var setPropsResult = _commands.SetProperties(
            workbookPath, 
            connectionName, 
            backgroundQuery: !originalBackgroundQuery, 
            refreshOnFileOpen: true,
            savePassword: null,
            refreshPeriod: null);
        Assert.True(setPropsResult.Success, $"SetProperties failed: {setPropsResult.ErrorMessage}");

        // Act 5: Verify property changes
        var verifyPropsResult = _commands.GetProperties(workbookPath, connectionName);
        Assert.True(verifyPropsResult.Success, $"GetProperties verification failed: {verifyPropsResult.ErrorMessage}");
        Assert.Equal(!originalBackgroundQuery, verifyPropsResult.BackgroundQuery);
        Assert.True(verifyPropsResult.RefreshOnFileOpen);

        // Act 6: Refresh connection
        var refreshResult = _commands.Refresh(workbookPath, connectionName);
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // Act 7: Test connection validity
        var testResult = _commands.Test(workbookPath, connectionName);
        Assert.True(testResult.Success, $"Test failed: {testResult.ErrorMessage}");
    }

    [Fact]
    public void MultipleConnections_CreateListUpdateDeleteAll_Success()
    {
        // Arrange: Create workbook with multiple connections of different types
        var workbookPath = CreateTestWorkbookWithMultipleConnections();
        var connection1 = "WebConnection1";
        var connection2 = "WebConnection2";
        var connection3 = "WebConnection3";

        // Act 1: List all connections
        var listResult = _commands.List(workbookPath);
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        var initialCount = listResult.Connections.Count;
        Assert.True(initialCount >= 3, "Expected at least 3 connections");

        // Act 2: Export each connection
        var export1Path = Path.Combine(_testDirectory, "conn1.json");
        var export2Path = Path.Combine(_testDirectory, "conn2.json");
        var export3Path = Path.Combine(_testDirectory, "conn3.json");

        var export1 = _commands.Export(workbookPath, connection1, export1Path);
        var export2 = _commands.Export(workbookPath, connection2, export2Path);
        var export3 = _commands.Export(workbookPath, connection3, export3Path);

        Assert.True(export1.Success && export2.Success && export3.Success, "One or more exports failed");
        Assert.True(File.Exists(export1Path) && File.Exists(export2Path) && File.Exists(export3Path));

        // Act 3: Test all connections
        var test1 = _commands.Test(workbookPath, connection1);
        var test2 = _commands.Test(workbookPath, connection2);
        var test3 = _commands.Test(workbookPath, connection3);

        Assert.True(test1.Success && test2.Success && test3.Success, "One or more tests failed");

        // Act 4: Get properties for all connections
        var props1 = _commands.GetProperties(workbookPath, connection1);
        var props2 = _commands.GetProperties(workbookPath, connection2);
        var props3 = _commands.GetProperties(workbookPath, connection3);

        Assert.True(props1.Success && props2.Success && props3.Success, "One or more property retrievals failed");

        // Act 5: Delete all connections
        var delete1 = _commands.Delete(workbookPath, connection1);
        var delete2 = _commands.Delete(workbookPath, connection2);
        var delete3 = _commands.Delete(workbookPath, connection3);

        Assert.True(delete1.Success && delete2.Success && delete3.Success, "One or more deletions failed");

        // Verify: List should show fewer connections
        var finalListResult = _commands.List(workbookPath);
        Assert.True(finalListResult.Success);
        Assert.True(finalListResult.Connections.Count < initialCount, "Connection count should decrease after deletions");
        Assert.DoesNotContain(finalListResult.Connections, c => c.Name == connection1);
        Assert.DoesNotContain(finalListResult.Connections, c => c.Name == connection2);
        Assert.DoesNotContain(finalListResult.Connections, c => c.Name == connection3);

        // Cleanup export files
        try
        {
            File.Delete(export1Path);
            File.Delete(export2Path);
            File.Delete(export3Path);
        }
        catch { /* Ignore cleanup errors */ }
    }

    #region Helper Methods

    private string CreateTestWorkbookWithConnection()
    {
        var filePath = Path.Combine(_testDirectory, $"workbook_with_connection_{Guid.NewGuid():N}.xlsx");
        _testFiles.Add(filePath);

        // Create workbook using FileCommands
        var fileCommands = new FileCommands();
        var createResult = fileCommands.CreateEmpty(filePath, overwriteIfExists: false);
        Assert.True(createResult.Success, $"Failed to create test workbook: {createResult.ErrorMessage}");

        // Create actual web connection for round trip testing
        ConnectionTestHelper.CreateWebConnection(filePath, "TestConnection", "https://example.com");

        return filePath;
    }

    private string CreateTestWorkbookWithTextConnection()
    {
        var filePath = Path.Combine(_testDirectory, $"workbook_text_conn_{Guid.NewGuid():N}.xlsx");
        _testFiles.Add(filePath);

        var fileCommands = new FileCommands();
        var createResult = fileCommands.CreateEmpty(filePath, overwriteIfExists: false);
        Assert.True(createResult.Success, $"Failed to create test workbook: {createResult.ErrorMessage}");

        // Create actual text file connection for round trip testing
        var textFilePath = Path.Combine(_testDirectory, $"test_data_{Guid.NewGuid():N}.csv");
        _testFiles.Add(textFilePath);
        File.WriteAllText(textFilePath, "Column1,Column2,Column3\nValue1,Value2,Value3\nData1,Data2,Data3");
        
        ConnectionTestHelper.CreateTextFileConnection(filePath, "DataConnection", textFilePath);
        
        return filePath;
    }

    private string CreateTestWorkbookWithMultipleConnections()
    {
        var filePath = Path.Combine(_testDirectory, $"workbook_multi_conn_{Guid.NewGuid():N}.xlsx");
        _testFiles.Add(filePath);

        var fileCommands = new FileCommands();
        var createResult = fileCommands.CreateEmpty(filePath, overwriteIfExists: false);
        Assert.True(createResult.Success, $"Failed to create test workbook: {createResult.ErrorMessage}");

        // Create multiple connections for round trip testing
        ConnectionTestHelper.CreateWebConnection(filePath, "WebConnection1", "https://example.com");
        ConnectionTestHelper.CreateWebConnection(filePath, "WebConnection2", "https://api.example.com");
        ConnectionTestHelper.CreateWebConnection(filePath, "WebConnection3", "https://data.example.com");
        
        return filePath;
    }

    #endregion
}
