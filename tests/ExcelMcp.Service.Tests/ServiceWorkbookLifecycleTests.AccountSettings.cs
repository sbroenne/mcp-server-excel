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
    [InlineData(true, true)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(false, false)]
    public async Task ConnectionProviderTransition_WithBackgroundChange_RejectsOnlyProviderChanges(bool fromOlap, bool providerChanged)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "provider-transition.xlsx"));
            sessions[session] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            batch.Execute((context, _) =>
            {
                Excel.Connections? connections = null;
                Excel.WorkbookConnection? connection = null;
                try
                {
                    connections = context.Book.Connections;
                    connection = connections.Add2("Selected", "Original description",
                        fromOlap
                            ? "OLEDB;Provider=MSOLAP.8;Data Source=localhost;Initial Catalog=Fixture;Connect Timeout=1;"
                            : "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=fixture.accdb;",
                        fromOlap ? "Model" : "SELECT 1", fromOlap ? Excel.XlCmdType.xlCmdCube : Excel.XlCmdType.xlCmdSql,
                        false, false);
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                    ComUtilities.Release(ref connections);
                }
            });
            var before = ReadAccountSettingsFixture(batch, "Selected");
            bool targetOlap = fromOlap != providerChanged;
            string requested = targetOlap
                ? "OLEDB;Provider=MSOLAP.8;Data Source=127.0.0.1;Initial Catalog=Fixture;Connect Timeout=1;"
                : "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=changed-fixture.accdb;";
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "connection.set-properties",
                SessionId = session,
                Args = JsonSerializer.Serialize(new
                {
                    connectionName = "Selected",
                    description = "Requested description",
                    connectionString = requested,
                    backgroundQuery = providerChanged || !targetOlap
                }, ServiceProtocol.JsonOptions)
            });
            if (providerChanged)
            {
                Assert.False(response.Success);
                Assert.Equal("InvalidInput", response.ErrorCategory);
                Assert.Contains("separate", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(before, ReadAccountSettingsFixture(batch, "Selected"));
            }
            else
            {
                RequireSuccess(response);
                var after = ReadAccountSettingsFixture(batch, "Selected");
                Assert.Equal("Requested description", after.Description);
                Assert.Equal(new DbConnectionStringBuilder { ConnectionString = requested[6..] }.ConnectionString,
                    new DbConnectionStringBuilder { ConnectionString = after.ConnectionString }.ConnectionString);
            }
        });
    }

    [Theory]
    [InlineData("accountHint", "UID")]
    [InlineData("interactiveLogin", "User ID")]
    [InlineData("identityMode", "User ID")]
    [InlineData("all", "User ID")]
    [InlineData("unchanged", "User ID")]
    public async Task ConnectionAccountSettings_SetOnlyRequestedFields_AndPersist(string selected, string hintKey)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "set-account-settings.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            AddAccountSettingsFixture(batch, "Selected", true, hintKey);
            AddAccountSettingsFixture(batch, "Selected neighbor", true);
            var before = ReadAccountSettingsFixture(batch, "Selected");
            var neighbor = ReadAccountSettingsFixture(batch, "Selected neighbor");
            string? hint = selected is "accountHint" or "all" ? "fixture;new-account" :
                selected == "unchanged" ? "fixture-account" : null;
            string? interactive = selected is "interactiveLogin" or "all" ? "Always" :
                selected == "unchanged" ? "Enabled" : null;
            string? identity = selected is "identityMode" or "all" ? "CurrentUser" :
                selected == "unchanged" ? "Connection" : null;
            var response = await SetAccountSettingsRequestAsync(service, session, hint, interactive, identity);
            RequireSuccess(response);
            using var json = JsonDocument.Parse(response.Result!);
            Assert.Equal(selected != "unchanged", json.RootElement.GetProperty("changed").GetBoolean());
            Assert.True(json.RootElement.GetProperty("accountHintPresent").GetBoolean());
            Assert.Equal(interactive ?? "Enabled", json.RootElement.GetProperty("interactiveLogin").GetString());
            Assert.Equal(identity ?? "Connection", json.RootElement.GetProperty("identityMode").GetString());
            Assert.DoesNotContain("fixture;new-account", response.Result!, StringComparison.Ordinal);
            Assert.DoesNotContain("fixture-secret", response.Result!, StringComparison.Ordinal);
            Assert.DoesNotContain("fixture-impersonation", response.Result!, StringComparison.Ordinal);
            var expected = new DbConnectionStringBuilder { ConnectionString = before.ConnectionString };
            if (hint != null)
            {
                expected.Remove("User ID");
                expected.Remove("UID");
                expected["User ID"] = hint;
            }
            if (interactive != null) expected["Interactive Login"] = interactive;
            if (identity != null) expected["Identity Mode"] = identity;
            AssertAccountSettingsMatch(expected, before, ReadAccountSettingsFixture(batch, "Selected"));
            Assert.Equal(neighbor, ReadAccountSettingsFixture(batch, "Selected neighbor"));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));

            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            session = await OpenSessionAsync(service, path);
            sessions[session] = 0;
            batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            AssertAccountSettingsMatch(expected, before, ReadAccountSettingsFixture(batch, "Selected"));
            Assert.Equal(neighbor, ReadAccountSettingsFixture(batch, "Selected neighbor"));
        });
    }

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
    [InlineData(null, null, null)]
    [InlineData("", null, null)]
    [InlineData(" ", "Always", null)]
    [InlineData("fixture\0account", null, null)]
    [InlineData("fixture-new-account", "Unsupported", null)]
    [InlineData("fixture-new-account", null, "Unsupported")]
    public async Task ConnectionAccountSettings_SetInvalidInput_PreservesWorkbook(
        string? hint, string? interactive, string? identity)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var session = await CreateSessionAsync(service, Path.Join(directory, "invalid-account-settings.xlsx"));
            sessions[session] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            AddAccountSettingsFixture(batch, "Selected", true);
            var before = ReadAccountSettingsFixture(batch, "Selected");
            var response = await SetAccountSettingsRequestAsync(service, session, hint, interactive, identity);
            Assert.False(response.Success);
            Assert.False(string.IsNullOrWhiteSpace(response.ErrorMessage));
            Assert.Equal(before, ReadAccountSettingsFixture(batch, "Selected"));
            Assert.DoesNotContain("fixture-secret", response.ErrorMessage, StringComparison.Ordinal);
        });
    }

    [Fact]
    public async Task ConnectionAccountSettings_SetRejectsReadOnlyWorkbook_ButInspectionSucceeds()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "readonly-account-settings.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            AddAccountSettingsFixture(batch, "Selected", true);
            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            session = await OpenSessionAsync(service, path);
            sessions[session] = 0;
            batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            batch.Execute((context, _) =>
                context.Book.ChangeFileAccess(Excel.XlFileAccess.xlReadOnly, Type.Missing, false));
            Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
            var before = ReadAccountSettingsFixture(batch, "Selected");
            RequireSuccess(await AccountSettingsRequestAsync(service, session, "get-account-settings", "Selected"));
            var response = await SetAccountSettingsRequestAsync(service, session, "fixture-new-account", null, null);
            Assert.False(response.Success);
            Assert.Contains("read-only", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before, ReadAccountSettingsFixture(batch, "Selected"));
            await CloseSessionAsync(service, session, save: false);
            sessions.TryRemove(session, out _);
        });
    }

    [Theory]
    [InlineData("get-account-settings")]
    [InlineData("clear-account-hint")]
    [InlineData("set-account-settings")]
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
            var response = action == "set-account-settings"
                ? await SetAccountSettingsRequestAsync(service, session, "fixture-account", null, null, name)
                : await AccountSettingsRequestAsync(service, session, action, name);
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

    private static Task<ServiceResponse> SetAccountSettingsRequestAsync(
        ExcelMcpService service, string session, string? accountHint, string? interactiveLogin,
        string? identityMode, string name = "Selected") =>
        service.ProcessAsync(new ServiceRequest
        {
            Command = "connection.set-account-settings",
            SessionId = session,
            Args = JsonSerializer.Serialize(new { connectionName = name, accountHint, interactiveLogin, identityMode }, ServiceProtocol.JsonOptions)
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
        AssertAccountSettingsMatch(expected, before, after);
    }

    private static void AssertAccountSettingsMatch(
        DbConnectionStringBuilder expected, AccountSettingsSnapshot before, AccountSettingsSnapshot after)
    {
        var actual = new DbConnectionStringBuilder { ConnectionString = after.ConnectionString };
        Assert.Equal(expected.Count, actual.Count);
        foreach (string key in expected.Keys)
            Assert.Equal(expected[key], actual[key]);
        Assert.Equal(before with { ConnectionString = after.ConnectionString }, after);
    }

    private sealed record AccountSettingsSnapshot(
        string ConnectionString, string Description, string? CommandText,
        bool SavePassword, bool RefreshWithRefreshAll, bool RefreshOnFileOpen);
}
