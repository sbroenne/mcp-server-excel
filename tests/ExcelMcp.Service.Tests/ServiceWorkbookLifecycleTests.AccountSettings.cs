using System.Data.Common;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(true, "User ID")]
    [InlineData(true, "UID")]
    [InlineData(true, "uSeR iD")]
    [InlineData(false, "User ID")]
    public async Task ConnectionAccountSettings_ClearOnlySelectedHint_AndPersist(bool hasHint, string hintKey)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "account-settings.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            AddAccountSettingsFixture(batch, "Selected", hasHint, hintKey);
            AddAccountSettingsFixture(batch, "Selected neighbor", true);
            var before = ReadAccountSettingsFixture(batch, "Selected");
            var neighbor = ReadAccountSettingsFixture(batch, "Selected neighbor");

            var inspected = await AccountSettingsRequestAsync(service, session, "get-account-settings", "Selected");
            RequireSuccess(inspected);
            using (var json = JsonDocument.Parse(inspected.Result!))
            {
                var result = json.RootElement;
                Assert.Equal(hasHint, result.GetProperty("accountHintPresent").GetBoolean());
                Assert.True(result.GetProperty("passwordPresent").GetBoolean());
                Assert.True(result.GetProperty("impersonationPresent").GetBoolean());
                Assert.Equal("Enabled", result.GetProperty("interactiveLogin").GetString());
                Assert.Equal("Connection", result.GetProperty("identityMode").GetString());
            }
            Assert.DoesNotContain("fixture-account", inspected.Result!, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("fixture-secret", inspected.Result!, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("fixture-impersonation", inspected.Result!, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before.ConnectionString, ReadAccountSettingsFixture(batch, "Selected").ConnectionString);

            var cleared = await AccountSettingsRequestAsync(service, session, "clear-account-hint", "Selected");
            RequireSuccess(cleared);
            using (var json = JsonDocument.Parse(cleared.Result!))
            {
                Assert.Equal(hasHint, json.RootElement.GetProperty("changed").GetBoolean());
                Assert.False(json.RootElement.GetProperty("accountHintPresent").GetBoolean());
            }
            Assert.DoesNotContain("fixture-secret", cleared.Result!, StringComparison.OrdinalIgnoreCase);
            var expected = new DbConnectionStringBuilder { ConnectionString = before.ConnectionString };
            expected.Remove("User ID");
            expected.Remove("UID");
            var after = ReadAccountSettingsFixture(batch, "Selected");
            AssertAccountSettingsPreserved(expected, before, after);
            if (!hasHint) Assert.Equal(before, after);
            Assert.Equal(neighbor, ReadAccountSettingsFixture(batch, "Selected neighbor"));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));

            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            session = await OpenSessionAsync(service, path);
            sessions[session] = 0;
            batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var persisted = ReadAccountSettingsFixture(batch, "Selected");
            AssertAccountSettingsPreserved(expected, before, persisted);
            Assert.Equal(neighbor, ReadAccountSettingsFixture(batch, "Selected neighbor"));
        });
    }

    [Theory]
    [InlineData("get-account-settings")]
    [InlineData("clear-account-hint")]
    public async Task ConnectionAccountSettings_RejectPowerQuery_WithoutChangingIt(string action)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "unsupported-account.xlsx"));
            sessions[session] = 0;
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = session,
                Args = """{"queryName":"AccountFixture","mCode":"#table({\"Value\"}, {{1}})","loadDestination":"worksheet","targetSheet":"AccountFixture"}"""
            }));
            var listed = await service.ProcessAsync(new ServiceRequest { Command = "connection.list", SessionId = session });
            RequireSuccess(listed);
            using var listing = JsonDocument.Parse(listed.Result!);
            var name = Assert.Single(listing.RootElement.GetProperty("connections").EnumerateArray()).GetProperty("name").GetString()!;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var before = ReadAccountSettingsFixture(batch, name);
            var response = await AccountSettingsRequestAsync(service, session, action, name);
            Assert.False(response.Success);
            Assert.Contains("MSOLAP", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(before, ReadAccountSettingsFixture(batch, name));
        });
    }

    private static Task<ServiceResponse> AccountSettingsRequestAsync(
        ExcelMcpService service, string session, string action, string name) =>
        service.ProcessAsync(new ServiceRequest
        {
            Command = "connection." + action,
            SessionId = session,
            Args = JsonSerializer.Serialize(new { connectionName = name }, ServiceProtocol.JsonOptions)
        });

    private static void AddAccountSettingsFixture(IExcelBatch batch, string name, bool hasHint, string hintKey = "User ID") =>
        batch.Execute((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Add2(name, "Account settings fixture",
                    "OLEDB;Provider=MSOLAP.8;Data Source=localhost;Initial Catalog=Fixture;Connect Timeout=1;" +
                    (hasHint ? hintKey + "=fixture-account;" : "") +
                    "Password=\"fixture-secret;with-semicolon\";EffectiveUserName=fixture-impersonation;" +
                    "Interactive Login=Enabled;Identity Mode=Connection;Persist Security Info=True;",
                    "Model", Excel.XlCmdType.xlCmdCube, false, false);
            }
            finally
            {
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });

    private static AccountSettingsSnapshot ReadAccountSettingsFixture(IExcelBatch batch, string name) =>
        batch.Execute((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledb = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Item(name);
                oledb = connection.OLEDBConnection;
                return new AccountSettingsSnapshot(
                    Convert.ToString(oledb.Connection, System.Globalization.CultureInfo.InvariantCulture)!.Replace("OLEDB;", "", StringComparison.OrdinalIgnoreCase),
                    connection.Description, Convert.ToString(oledb.CommandText, System.Globalization.CultureInfo.InvariantCulture),
                    oledb.SavePassword, connection.RefreshWithRefreshAll, oledb.RefreshOnFileOpen);
            }
            finally
            {
                ComUtilities.Release(ref oledb);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });

    private static void AssertAccountSettingsPreserved(
        DbConnectionStringBuilder expected, AccountSettingsSnapshot before, AccountSettingsSnapshot after)
    {
        var actual = new DbConnectionStringBuilder { ConnectionString = after.ConnectionString };
        Assert.False(actual.ContainsKey("User ID"));
        Assert.False(actual.ContainsKey("UID"));
        Assert.Equal(expected.Count, actual.Count);
        foreach (string key in expected.Keys)
            Assert.Equal(expected[key], actual[key]);
        Assert.Equal(before with { ConnectionString = after.ConnectionString }, after);
    }

    private sealed record AccountSettingsSnapshot(
        string ConnectionString, string Description, string? CommandText,
        bool SavePassword, bool RefreshWithRefreshAll, bool RefreshOnFileOpen);
}
