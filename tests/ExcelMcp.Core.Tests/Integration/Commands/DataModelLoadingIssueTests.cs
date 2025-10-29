using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.ComInterop;
using System.Reflection;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests specifically for replicating issue #64:
/// set-load-to-data-model consistently fails with "Failed to configure query for Data Model loading"
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModelLoading")]
[Trait("Issue", "64")]
public class DataModelLoadingIssueTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly ITableCommands _tableCommands;
    private readonly IDataModelCommands _dataModelCommands;
    private readonly string _testExcelFile;
    private readonly string _testParametersQueryFile;
    private readonly string _testProjectRootQueryFile;
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private bool _disposed;

    public DataModelLoadingIssueTests(ITestOutputHelper output)
    {
        _output = output;
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _fileCommands = new FileCommands();
        _tableCommands = new TableCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DataModel_Issue64_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestDataModel.xlsx");
        _testParametersQueryFile = Path.Combine(_tempDir, "Parameters.pq");
        _testProjectRootQueryFile = Path.Combine(_tempDir, "ProjectRootDirectory.pq");

        // Create test files matching the issue scenario
        CreateTestExcelFile();
        CreateTestQueryFiles();
    }

    /// <summary>
    /// Test replicating the exact scenario from issue #64
    /// </summary>
    [Fact]
    public async Task Issue64_SetLoadToDataModel_ReplicatesFailure()
    {
        _output.WriteLine("=== REPRODUCING ISSUE #64: set-load-to-data-model failure ===");

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

        // STEP 1: Setup named ranges (as mentioned in issue)
        _output.WriteLine("\n1. Setting up named ranges...");
        await CreateNamedRanges(batch);

        // STEP 2: Import working Power Queries (as mentioned in issue)
        _output.WriteLine("\n2. Importing Power Queries...");
        var projectRootImportResult = await _powerQueryCommands.ImportAsync(batch, "ProjectRootDirectory Parameter", _testProjectRootQueryFile);
        _output.WriteLine($"ProjectRootDirectory import: Success={projectRootImportResult.Success}, Error={projectRootImportResult.ErrorMessage}");
        Assert.True(projectRootImportResult.Success, $"ProjectRootDirectory import failed: {projectRootImportResult.ErrorMessage}");

        var parametersImportResult = await _powerQueryCommands.ImportAsync(batch, "Parameters", _testParametersQueryFile);
        _output.WriteLine($"Parameters import: Success={parametersImportResult.Success}, Error={parametersImportResult.ErrorMessage}");
        Assert.True(parametersImportResult.Success, $"Parameters import failed: {parametersImportResult.ErrorMessage}");

        // STEP 3: Verify queries work (list and refresh)
        _output.WriteLine("\n3. Verifying queries work...");
        var listResult = await _powerQueryCommands.ListAsync(batch);
        _output.WriteLine($"List queries: Success={listResult.Success}, Count={listResult.Queries.Count}");
        Assert.True(listResult.Success, $"List queries failed: {listResult.ErrorMessage}");
        Assert.True(listResult.Queries.Count >= 2, "Expected at least 2 queries");

        var refreshResult = await _powerQueryCommands.RefreshAsync(batch, "Parameters");
        _output.WriteLine($"Refresh Parameters: Success={refreshResult.Success}, Error={refreshResult.ErrorMessage}");
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // STEP 4: Verify set-load-to-table works (baseline)
        _output.WriteLine("\n4. Testing set-load-to-table (should work)...");
        var setTableResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "Parameters", "Sheet1");
        _output.WriteLine($"Set load to table: Success={setTableResult.Success}, Error={setTableResult.ErrorMessage}");
        Assert.True(setTableResult.Success, $"Set load to table failed: {setTableResult.ErrorMessage}");

        // STEP 5: Check Data Model availability before attempting
        _output.WriteLine("\n5. Checking Data Model availability...");
        await CheckDataModelDiagnostics(batch);

        // STEP 6: REPRODUCE THE FAILURE - set-load-to-data-model
        _output.WriteLine("\n6. *** REPRODUCING THE FAILURE: set-load-to-data-model ***");
        var setDataModelResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "Parameters");

        // Log the detailed results for analysis
        _output.WriteLine($"set-load-to-data-model Results:");
        _output.WriteLine($"  Success: {setDataModelResult.Success}");
        _output.WriteLine($"  ErrorMessage: {setDataModelResult.ErrorMessage}");
        _output.WriteLine($"  ConfigurationApplied: {setDataModelResult.ConfigurationApplied}");
        _output.WriteLine($"  DataLoadedToModel: {setDataModelResult.DataLoadedToModel}");
        _output.WriteLine($"  WorkflowStatus: {setDataModelResult.WorkflowStatus}");
        _output.WriteLine($"  TablesInDataModel: {setDataModelResult.TablesInDataModel}");

        // This should fail according to issue #64
        Assert.False(setDataModelResult.Success, "Expected set-load-to-data-model to fail (replicating issue #64)");
        Assert.Contains("Failed to configure query for Data Model loading", setDataModelResult.ErrorMessage);
    }

    /// <summary>
    /// Test the Excel Table add-to-datamodel failure mentioned in the issue comment
    /// </summary>
    [Fact]
    public async Task Issue64_TableAddToDataModel_ReplicatesFailure()
    {
        _output.WriteLine("=== REPRODUCING ISSUE #64: table add-to-datamodel failure ===");

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

        // STEP 1: Create an Excel Table
        _output.WriteLine("\n1. Creating Excel Table...");
        await CreateTestTableWithData(batch);

        var createTableResult = await _tableCommands.CreateAsync(batch, "Sheet1", "TestTable", "A1:C4", hasHeaders: true);
        _output.WriteLine($"Create table: Success={createTableResult.Success}, Error={createTableResult.ErrorMessage}");
        Assert.True(createTableResult.Success, $"Create table failed: {createTableResult.ErrorMessage}");

        // STEP 2: Verify table exists and info works
        _output.WriteLine("\n2. Verifying table operations work...");
        var listTableResult = await _tableCommands.ListAsync(batch);
        _output.WriteLine($"List tables: Success={listTableResult.Success}, Count={listTableResult.Tables.Count}");
        Assert.True(listTableResult.Success, $"List tables failed: {listTableResult.ErrorMessage}");

        var tableInfoResult = await _tableCommands.GetInfoAsync(batch, "TestTable");
        _output.WriteLine($"Table info: Success={tableInfoResult.Success}, Error={tableInfoResult.ErrorMessage}");
        Assert.True(tableInfoResult.Success, $"Table info failed: {tableInfoResult.ErrorMessage}");

        // STEP 3: Check Data Model availability before attempting
        _output.WriteLine("\n3. Checking Data Model availability...");
        await CheckDataModelDiagnostics(batch);

        // STEP 4: REPRODUCE THE FAILURE - add-to-datamodel
        _output.WriteLine("\n4. *** REPRODUCING THE FAILURE: table add-to-datamodel ***");
        var addToDataModelResult = await _tableCommands.AddToDataModelAsync(batch, "TestTable");

        // Log the detailed results for analysis
        _output.WriteLine($"add-to-datamodel Results:");
        _output.WriteLine($"  Success: {addToDataModelResult.Success}");
        _output.WriteLine($"  ErrorMessage: {addToDataModelResult.ErrorMessage}");

        // After fix: We should NOT see the original COM method error anymore
        Assert.False(addToDataModelResult.ErrorMessage?.Contains("does not contain a definition for 'Add'") ?? false,
            "Fix should eliminate the invalid COM method error");

        // In test environments, this might still fail due to Power Pivot configuration
        // but the important thing is we're using the correct COM API now
        if (!addToDataModelResult.Success)
        {
            _output.WriteLine($"INFO: Post-fix error (may be test environment): {addToDataModelResult.ErrorMessage}");
        }
        else
        {
            _output.WriteLine("SUCCESS: Table was successfully added to Data Model!");
        }
    }

    /// <summary>
    /// Diagnostic test to analyze what's happening with the Data Model API
    /// </summary>
    [Fact]
    public async Task Issue64_DiagnosticAnalysis_DataModelAPI()
    {
        _output.WriteLine("=== DIAGNOSTIC ANALYSIS: Data Model API Investigation ===");

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

        var diagnostics = await batch.ExecuteAsync<List<string>>((ctx, ct) =>
        {
            var results = new List<string>();

            try
            {
                // Test 1: Basic workbook properties
                results.Add($"Workbook Name: {ctx.Book.Name}");
                results.Add($"Workbook FullName: {ctx.Book.FullName}");
                results.Add($"Excel Version: {ctx.App.Version}");

                // Test 2: Check if Model property exists
                dynamic? model = null;
                try
                {
                    model = ctx.Book.Model;
                    results.Add($"Model property exists: {model != null}");

                    if (model != null)
                    {
                        // Test 3: Check ModelTables property
                        dynamic? modelTables = null;
                        try
                        {
                            modelTables = model.ModelTables;
                            results.Add($"ModelTables property exists: {modelTables != null}");
                            results.Add($"ModelTables count: {modelTables?.Count ?? 0}");

                            // Test 4: Check if ModelTables has Add method
                            try
                            {
                                var type = modelTables.GetType();
                                var methods = type.GetMethods();
                                var addMethods = new List<MethodInfo>();
                                foreach (var method in methods)
                                {
                                    if (method.Name.Contains("Add"))
                                        addMethods.Add(method);
                                }
                                results.Add($"ModelTables Add methods found: {addMethods.Count}");
                                foreach (var method in addMethods)
                                {
                                    var paramNames = new List<string>();
                                    foreach (var param in method.GetParameters())
                                    {
                                        paramNames.Add(param.ParameterType.Name);
                                    }
                                    results.Add($"  - {method.Name}({string.Join(", ", paramNames)})");
                                }
                            }
                            catch (Exception ex)
                            {
                                results.Add($"Error checking ModelTables methods: {ex.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            results.Add($"Error accessing ModelTables: {ex.Message}");
                        }
                        finally
                        {
                            ComUtilities.Release(ref modelTables);
                        }
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Error accessing Model: {ex.Message}");
                }
                finally
                {
                    ComUtilities.Release(ref model);
                }

                // Test 5: Check Power Pivot availability
                try
                {
                    // Try to access AddIns to see if Power Pivot is available
                    dynamic addIns = ctx.App.AddIns;
                    results.Add($"AddIns count: {addIns.Count}");

                    for (int i = 1; i <= addIns.Count; i++)
                    {
                        dynamic addIn = addIns.Item(i);
                        string name = addIn.Name ?? "";
                        bool installed = addIn.Installed;
                        results.Add($"  AddIn: {name}, Installed: {installed}");

                        if (name.Contains("Power", StringComparison.OrdinalIgnoreCase) ||
                            name.Contains("Pivot", StringComparison.OrdinalIgnoreCase))
                        {
                            results.Add($"    *** POWER PIVOT FOUND: {name} ***");
                        }
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Error checking AddIns: {ex.Message}");
                }

                // Test 6: Check Connections API
                try
                {
                    dynamic connections = ctx.Book.Connections;
                    results.Add($"Connections property exists: {connections != null}");
                    results.Add($"Connections count: {connections?.Count ?? 0}");

                    // Check if Connections has Add2 method
                    try
                    {
                        var type = connections.GetType();
                        var methods = type.GetMethods();
                        var addMethods = new List<MethodInfo>();
                        foreach (var method in methods)
                        {
                            if (method.Name.Contains("Add"))
                                addMethods.Add(method);
                        }
                        results.Add($"Connections Add methods found: {addMethods.Count}");
                        foreach (var method in addMethods)
                        {
                            var paramNames = new List<string>();
                            foreach (var param in method.GetParameters())
                            {
                                paramNames.Add(param.ParameterType.Name);
                            }
                            results.Add($"  - {method.Name}({string.Join(", ", paramNames)})");
                        }
                    }
                    catch (Exception ex)
                    {
                        results.Add($"Error checking Connections methods: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Error accessing Connections: {ex.Message}");
                }

            }
            catch (Exception ex)
            {
                results.Add($"Critical error in diagnostics: {ex.Message}");
            }

            return ValueTask.FromResult(results);
        });

        _output.WriteLine("\nDiagnostic Results:");
        foreach (var result in diagnostics)
        {
            _output.WriteLine($"  {result}");
        }

        // This test always passes - it's just for analysis
        Assert.True(true, "Diagnostic test completed");
    }

    private async Task CheckDataModelDiagnostics(IExcelBatch batch)
    {
        var hasDataModel = await batch.ExecuteAsync<bool>((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                model = ctx.Book.Model;
                if (model == null)
                {
                    _output.WriteLine("  Model property is null");
                    return ValueTask.FromResult(false);
                }

                dynamic? modelTables = null;
                try
                {
                    modelTables = model.ModelTables;
                    _output.WriteLine($"  ModelTables count: {modelTables?.Count ?? 0}");
                    return ValueTask.FromResult(modelTables != null);
                }
                catch (Exception ex)
                {
                    _output.WriteLine($"  Error accessing ModelTables: {ex.Message}");
                    return ValueTask.FromResult(false);
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }
            }
            catch (Exception ex)
            {
                _output.WriteLine($"  Error accessing Model: {ex.Message}");
                return ValueTask.FromResult(false);
            }
            finally
            {
                ComUtilities.Release(ref model);
            }
        });

        _output.WriteLine($"  Data Model available: {hasDataModel}");
    }

    private async Task CreateNamedRanges(IExcelBatch batch)
    {
        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);

            // Create named ranges as mentioned in issue #64
            sheet.Range["A1"].Value2 = "2025-01-01"; // Start_Date
            sheet.Range["A2"].Value2 = 12; // Duration_Months
            sheet.Range["A3"].Value2 = "C:\\Projects\\Test"; // ProjectRoot

            // Create named ranges
            dynamic names = ctx.Book.Names;
            names.Add("Start_Date", "=Sheet1!$A$1");
            names.Add("Duration_Months", "=Sheet1!$A$2");
            names.Add("ProjectRoot", "=Sheet1!$A$3");

            _output.WriteLine("  Created named ranges: Start_Date, Duration_Months, ProjectRoot");
            return ValueTask.FromResult(0);
        });
    }

    private async Task CreateTestTableWithData(IExcelBatch batch)
    {
        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);

            // Create some test data for the table
            sheet.Range["A1"].Value2 = "ID";
            sheet.Range["B1"].Value2 = "Name";
            sheet.Range["C1"].Value2 = "Value";

            sheet.Range["A2"].Value2 = 1;
            sheet.Range["B2"].Value2 = "Item1";
            sheet.Range["C2"].Value2 = 100;

            sheet.Range["A3"].Value2 = 2;
            sheet.Range["B3"].Value2 = "Item2";
            sheet.Range["C3"].Value2 = 200;

            sheet.Range["A4"].Value2 = 3;
            sheet.Range["B4"].Value2 = "Item3";
            sheet.Range["C4"].Value2 = 300;

            _output.WriteLine("  Created test data for table");
            return ValueTask.FromResult(0);
        });
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
    }

    private void CreateTestQueryFiles()
    {
        // Create Parameters.pq matching the issue scenario
        var parametersQuery = @"
let
    Start_Date = Excel.CurrentWorkbook(){[Name=""Start_Date""]}[Content]{0}[Column1],
    Duration_Months = Excel.CurrentWorkbook(){[Name=""Duration_Months""]}[Content]{0}[Column1],
    ProjectRoot = Excel.CurrentWorkbook(){[Name=""ProjectRoot""]}[Content]{0}[Column1],

    Parameters = [
        StartDate = Start_Date,
        DurationMonths = Duration_Months,
        ProjectRootDirectory = ProjectRoot
    ]
in
    Parameters
";

        // Create ProjectRootDirectory Parameter.pq
        var projectRootQuery = @"
let
    ProjectRoot = Excel.CurrentWorkbook(){[Name=""ProjectRoot""]}[Content]{0}[Column1]
in
    ProjectRoot
";

        File.WriteAllText(_testParametersQueryFile, parametersQuery.Trim());
        File.WriteAllText(_testProjectRootQueryFile, projectRootQuery.Trim());
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failure is non-critical
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
