using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using System.Runtime.InteropServices;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Diagnostics;

/// <summary>
/// Diagnostic test to pinpoint exact location of 0x800A03EC exception in PowerQuery List()
/// </summary>
[Trait("RunType", "OnDemand")]
[Trait("Layer", "Diagnostics")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryListDiagnostic
{
    private readonly ITestOutputHelper _output;

    public PowerQueryListDiagnostic(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void DiagnoseConsumptionPlanBaseException()
    {
        const string testFile = @"D:\source\mcp-server-excel\ConsumptionPlan_Base.xlsx";

        if (!System.IO.File.Exists(testFile))
        {
            _output.WriteLine("Test file not found - skipping");
            return;
        }

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;
                _output.WriteLine($"Total queries: {count}");

                for (int i = 1; i <= count; i++)
                {
                    dynamic? query = null;
                    try
                    {
                        _output.WriteLine($"\n=== Processing Query {i} ===");

                        query = queriesCollection.Item(i);
                        _output.WriteLine($"✓ Got query object");

                        string name = "UNKNOWN";
                        try
                        {
                            name = query.Name ?? $"Query{i}";
                            _output.WriteLine($"✓ Name: {name}");
                        }
                        catch (COMException ex)
                        {
                            _output.WriteLine($"✗ Name access failed: 0x{ex.HResult:X} - {ex.Message}");
                            throw;
                        }

                        try
                        {
                            string formula = query.Formula?.ToString() ?? "";
                            _output.WriteLine($"✓ Formula length: {formula.Length}");
                        }
                        catch (COMException ex)
                        {
                            _output.WriteLine($"✗ Formula access failed: 0x{ex.HResult:X} - {ex.Message}");
                        }

                        // Check IsConnectionOnly logic
                        _output.WriteLine("Checking IsConnectionOnly...");
                        dynamic? worksheets = null;
                        try
                        {
                            worksheets = ctx.Book.Worksheets;
                            _output.WriteLine($"✓ Got worksheets: {worksheets.Count}");

                            for (int ws = 1; ws <= worksheets.Count; ws++)
                            {
                                dynamic? worksheet = null;
                                dynamic? listObjects = null;
                                try
                                {
                                    worksheet = worksheets.Item(ws);
                                    string sheetName = worksheet.Name;
                                    listObjects = worksheet.ListObjects;
                                    _output.WriteLine($"  Sheet {ws} ({sheetName}): {listObjects.Count} ListObjects");

                                    for (int lo = 1; lo <= listObjects.Count; lo++)
                                    {
                                        dynamic? listObject = null;
                                        dynamic? queryTable = null;
                                        dynamic? wbConn = null;
                                        dynamic? oledbConn = null;
                                        try
                                        {
                                            listObject = listObjects.Item(lo);
                                            queryTable = listObject.QueryTable;
                                            if (queryTable == null)
                                            {
                                                _output.WriteLine($"    ListObject {lo}: No QueryTable");
                                                continue;
                                            }

                                            wbConn = queryTable.WorkbookConnection;
                                            if (wbConn == null)
                                            {
                                                _output.WriteLine($"    ListObject {lo}: No WorkbookConnection");
                                                continue;
                                            }

                                            oledbConn = wbConn.OLEDBConnection;
                                            if (oledbConn == null)
                                            {
                                                _output.WriteLine($"    ListObject {lo}: No OLEDBConnection");
                                                continue;
                                            }

                                            string connString = oledbConn.Connection?.ToString() ?? "";
                                            bool isMashup = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                            bool locationMatches = connString.Contains($"Location={name}", StringComparison.OrdinalIgnoreCase);
                                            _output.WriteLine($"    ListObject {lo}: Mashup={isMashup}, LocationMatches={locationMatches}");
                                        }
                                        catch (COMException ex)
                                        {
                                            _output.WriteLine($"    ListObject {lo}: EXCEPTION 0x{ex.HResult:X} - {ex.Message}");
                                            throw;
                                        }
                                        finally
                                        {
                                            if (oledbConn != null) ComUtilities.Release(ref oledbConn!);
                                            if (wbConn != null) ComUtilities.Release(ref wbConn!);
                                            if (queryTable != null) ComUtilities.Release(ref queryTable!);
                                            if (listObject != null) ComUtilities.Release(ref listObject!);
                                        }
                                    }
                                }
                                catch (COMException ex)
                                {
                                    _output.WriteLine($"  Sheet {ws}: EXCEPTION 0x{ex.HResult:X} - {ex.Message}");
                                    throw;
                                }
                                finally
                                {
                                    if (listObjects != null) ComUtilities.Release(ref listObjects!);
                                    if (worksheet != null) ComUtilities.Release(ref worksheet!);
                                }
                            }
                        }
                        finally
                        {
                            if (worksheets != null) ComUtilities.Release(ref worksheets!);
                        }

                        _output.WriteLine($"✓ Query {i} processed successfully");
                    }
                    catch (Exception ex)
                    {
                        _output.WriteLine($"✗✗✗ CAUGHT EXCEPTION for Query {i}: {ex.GetType().Name} - {ex.Message}");
                        if (ex is COMException comEx)
                        {
                            _output.WriteLine($"    HResult: 0x{comEx.HResult:X}");
                        }
                        _output.WriteLine($"    Stack: {ex.StackTrace}");
                    }
                    finally
                    {
                        if (query != null) ComUtilities.Release(ref query!);
                    }
                }
            }
            finally
            {
                if (queriesCollection != null) ComUtilities.Release(ref queriesCollection!);
            }
        });
    }
}
