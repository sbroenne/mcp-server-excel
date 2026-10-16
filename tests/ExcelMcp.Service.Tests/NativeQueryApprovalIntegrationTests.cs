using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Settings;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
[Trait("RunType", "OnDemand")]
public sealed class NativeQueryApprovalIntegrationTests(
    NativeQueryApprovalWorkbookFixture fixture) :
    IClassFixture<NativeQueryApprovalWorkbookFixture>
{
    [Fact]
    public async Task Create_NativeQueryApprovalPrompt_ReportsPendingAndPreservesSettings()
    {
        var settings = new SettingsCommands();
        var settingsBefore = settings.GetNativeQueryApproval();
        Assert.True(settingsBefore.Required,
            "This manual test requires native-query approval to already be enabled. It does not change the setting.");

        var scope = fixture.CreateScope();
        try
        {
            var queryName = $"NativeApproval_{Guid.NewGuid():N}"[..24];
            scope.RegisterPowerQueryForCleanup(queryName);
            scope.RegisterSheetForCleanup(queryName);

            const string sql = "SELECT * FROM [Products$] WHERE [Price] > 15.0000";
            var connection = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={fixture.SourceWorkbookPath};Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            var mCode = $"let Source = OleDb.Query({ToMString(connection)}, {ToMString(sql)}) in Source";

            var queries = scope.CreateCommands<IPowerQueryCommands>();
            var createTask = Task.Run(() => queries.Create(
                scope.BatchToken,
                queryName,
                mCode,
                PowerQueryLoadMode.LoadToTable,
                queryName));

            var detectionDeadline = Stopwatch.StartNew();
            JsonElement session = default;
            while (detectionDeadline.Elapsed < TimeSpan.FromSeconds(10))
            {
                var response = await fixture.SendAsync("session.list", new { });
                Assert.True(response.Success, response.ErrorMessage);
                using var document = JsonDocument.Parse(response.Result!);
                session = Assert.Single(document.RootElement.GetProperty("sessions").EnumerateArray()).Clone();
                if (session.GetProperty("excelState").GetString() == "nativeQueryApproval")
                {
                    break;
                }
                await Task.Delay(100);
            }

            Assert.Equal("nativeQueryApproval", session.GetProperty("excelState").GetString());
            Assert.Equal("nativeQueryApproval", session.GetProperty("blockingReason").GetString());
            Assert.True(session.GetProperty("operationPending").GetBoolean());
            Assert.False(session.GetProperty("canClose").GetBoolean());
            Assert.Contains("approve the query manually",
                session.GetProperty("blockingMessage").GetString(),
                StringComparison.Ordinal);
            Assert.DoesNotContain(sql,
                session.GetProperty("blockingMessage").GetString(),
                StringComparison.Ordinal);

            var result = await createTask.WaitAsync(TimeSpan.FromMinutes(3));
            Assert.True(result.Success, result.ErrorMessage);
            PowerQueryStateAssertions.AssertStored(
                scope,
                queryName,
                mCode,
                PowerQueryLoadMode.LoadToTable,
                queryName,
                ["Product", "Price"],
                [["Widget", 19.99], ["Gadget", 29.99]]);

            var ready = await fixture.SendAsync("session.list", new { });
            using var readyDocument = JsonDocument.Parse(ready.Result!);
            var readySession = Assert.Single(readyDocument.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.False(readySession.GetProperty("operationPending").GetBoolean());
            Assert.Equal(settingsBefore.Required, settings.GetNativeQueryApproval().Required);
            Assert.Equal(settingsBefore.ExplicitlySet, settings.GetNativeQueryApproval().ExplicitlySet);
        }
        finally
        {
            await scope.DisposeAsync();
        }
    }

    private static string ToMString(string value) =>
        $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
}

public sealed class NativeQueryApprovalWorkbookFixture : PersistentServiceWorkbookFixture
{
    public NativeQueryApprovalWorkbookFixture()
        : base(CreateTestFiles, "NativeQueryApproval.xlsx", show: true)
    {
    }

    internal string SourceWorkbookPath =>
        Path.Combine(Path.GetDirectoryName(WorkbookPath)!, "NativeApprovalSource.xlsx");

    private static void CreateTestFiles(string workbookPath)
    {
        SavedWorkbookTemplates.CopyBlankTo(workbookPath);
        AceOleDbTestHelper.CreateExcelDataSource(
            Path.Combine(Path.GetDirectoryName(workbookPath)!, "NativeApprovalSource.xlsx"));
    }
}
