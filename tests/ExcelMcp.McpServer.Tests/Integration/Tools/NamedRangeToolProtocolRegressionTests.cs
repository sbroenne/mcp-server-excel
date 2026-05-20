// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end regressions for named range tool behavior through the MCP protocol.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public sealed class NamedRangeToolProtocolRegressionTests : McpIntegrationTestBase
{
    private readonly string _tempDir;

    public NamedRangeToolProtocolRegressionTests(ITestOutputHelper output)
        : base(output, "NamedRangeToolProtocolRegressionClient")
    {
        _tempDir = CreateTempDirectory("NamedRangeToolProtocolRegressionTests");
    }

    [Fact]
    public async Task NamedRangeList_EmptyWorkbook_ReturnsStructuredEmptyList_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"NoNamedRanges_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var listResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"namedrange list failed: {listResult}");
            Assert.True(root.TryGetProperty("namedRanges", out var namedRanges), $"namedrange list should return namedRanges: {listResult}");
            Assert.Equal(JsonValueKind.Array, namedRanges.ValueKind);
            Assert.Empty(namedRanges.EnumerateArray());
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after namedrange list");

        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task NamedRangeList_WorkbookWithNamedRange_ReturnsStructuredList_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"OneNamedRange_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        var name = $"CsvFolder_{Guid.NewGuid():N}";

        var createResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["name"] = name,
            ["reference"] = "Sheet1!$B$4"
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(createResult, "namedrange create setup");

        var writeResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "write",
            ["session_id"] = sessionId,
            ["name"] = name,
            ["value"] = "C:\\Data"
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(writeResult, "namedrange write setup");

        var listResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"namedrange list failed: {listResult}");
            var namedRanges = root.GetProperty("namedRanges").EnumerateArray().ToList();
            var listedRange = Assert.Single(namedRanges, range => range.GetProperty("name").GetString() == name);
            Assert.Contains("$B$4", listedRange.GetProperty("refersTo").GetString(), StringComparison.OrdinalIgnoreCase);
            Assert.Equal("C:\\Data", listedRange.GetProperty("value").GetString());
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after namedrange list");

        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task NamedRangeList_WorkbookWithHiddenName_ReturnsEmptyList_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"HiddenName_{Guid.NewGuid():N}.xlsx");
        var name = $"HiddenName_{Guid.NewGuid():N}";
        CreateWorkbookWithHiddenName(workbookPath, name);
        var sessionId = await OpenWorkbookSessionAsync(workbookPath);

        var listResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"namedrange list failed: {listResult}");
            var namedRanges = root.GetProperty("namedRanges").EnumerateArray().ToList();
            Assert.Empty(namedRanges);
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after namedrange list");

        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task NamedRangeList_LargeNamedRange_OmitsValuePreview_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"LargeNamedRange_{Guid.NewGuid():N}.xlsx");
        var name = $"LargePreview_{Guid.NewGuid():N}";
        CreateWorkbookWithNamedRange(workbookPath, name, "Sheet1!$A$1:$A$10001");
        var sessionId = await OpenWorkbookSessionAsync(workbookPath);

        var listResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"namedrange list failed: {listResult}");
            var namedRanges = root.GetProperty("namedRanges").EnumerateArray().ToList();
            var listedRange = Assert.Single(namedRanges, range => range.GetProperty("name").GetString() == name);
            Assert.Equal("RangeTooLarge", listedRange.GetProperty("valueType").GetString());
            Assert.False(listedRange.TryGetProperty("value", out _), $"large named range list should omit value: {listResult}");
            Assert.Equal(10001, listedRange.GetProperty("cellCount").GetInt64());
            Assert.Contains(
                "exceeds",
                listedRange.GetProperty("valueOmittedReason").GetString(),
                StringComparison.OrdinalIgnoreCase);
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after namedrange list");

        await CloseSessionAsync(sessionId, save: false);
    }

    private async Task<string> OpenWorkbookSessionAsync(string workbookPath)
    {
        var openJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = workbookPath
        });

        AssertSetupSuccess(openJson, $"file.open ({Path.GetFileName(workbookPath)})");

        using var openDoc = ParseJsonResult(openJson, $"file.open ({Path.GetFileName(workbookPath)})");
        var sessionId = openDoc.RootElement.GetProperty("session_id").GetString();
        TrackSession(sessionId);
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    private static void CreateWorkbookWithHiddenName(string workbookPath, string name)
    {
        CreateWorkbook(workbookPath, batch =>
        {
            batch.Execute((ctx, ct) =>
            {
                dynamic? names = null;
                dynamic? nameObj = null;
                try
                {
                    names = ctx.Book.Names;
                    nameObj = names.Add(name, "=Sheet1!$B$4");
                    nameObj.Visible = false;
                    return 0;
                }
                finally
                {
                    ComUtilities.Release(ref nameObj);
                    ComUtilities.Release(ref names);
                }
            });
        });
    }

    private static void CreateWorkbookWithNamedRange(string workbookPath, string name, string reference)
    {
        CreateWorkbook(workbookPath, batch =>
        {
            batch.Execute((ctx, ct) =>
            {
                dynamic? names = null;
                dynamic? nameObj = null;
                try
                {
                    names = ctx.Book.Names;
                    nameObj = names.Add(name, $"={reference.TrimStart('=')}");
                    return 0;
                }
                finally
                {
                    ComUtilities.Release(ref nameObj);
                    ComUtilities.Release(ref names);
                }
            });
        });
    }

    private static void CreateWorkbook(string workbookPath, Action<IExcelBatch> configure)
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(workbookPath, show: false);
        manager.CloseSession(sessionId, save: true);

        using var batch = ExcelSession.BeginBatch(workbookPath);
        configure(batch);
        batch.Save();
    }
}
