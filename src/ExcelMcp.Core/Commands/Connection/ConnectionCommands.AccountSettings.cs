using System.Data.Common;
using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class ConnectionCommands
{
    /// <inheritdoc />
    public ConnectionAccountSettingsResult GetAccountSettings(IExcelBatch batch, string connectionName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(connectionName);
        ValidateAccountSettingsReadiness(batch);
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledb = null;
            try
            {
                connection = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName)
                    ?? throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                oledb = GetAccountSettingsConnection(connection);
                var settings = ReadAccountSettingsConnectionString(oledb);
                ct.ThrowIfCancellationRequested();
                var result = new ConnectionAccountSettingsResult();
                PopulateAccountSettingsResult(result, batch, connectionName, oledb, settings);
                return result;
            }
            finally
            {
                ComUtilities.Release(ref oledb);
                ComUtilities.Release(ref connection);
            }
        });
    }

    /// <inheritdoc />
    public ConnectionAccountSettingsUpdateResult SetAccountSettings(
        IExcelBatch batch, string connectionName, string? accountHint = null,
        ConnectionInteractiveLogin? interactiveLogin = null, ConnectionIdentityMode? identityMode = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(connectionName);
        ValidateAccountSettingsUpdate(accountHint, interactiveLogin, identityMode);
        ValidateAccountSettingsReadiness(batch);
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledb = null;
            try
            {
                connection = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName)
                    ?? throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                oledb = GetAccountSettingsConnection(connection);
                var settings = ReadAccountSettingsConnectionString(oledb);
                var expected = new DbConnectionStringBuilder { ConnectionString = settings.ConnectionString };
                if (accountHint != null)
                {
                    expected.Remove("User ID");
                    expected.Remove("UID");
                    expected["User ID"] = accountHint;
                }
                if (interactiveLogin.HasValue) expected["Interactive Login"] = interactiveLogin.Value.ToString();
                if (identityMode.HasValue) expected["Identity Mode"] = identityMode.Value.ToString();
                bool changed = !AccountSettingsMatch(settings, expected);
                ct.ThrowIfCancellationRequested();
                if (changed) WriteAccountSettingsAndVerify(oledb, expected);
                var result = new ConnectionAccountSettingsUpdateResult { Changed = changed };
                PopulateAccountSettingsResult(result, batch, connectionName, oledb, expected);
                return result;
            }
            finally
            {
                ComUtilities.Release(ref oledb);
                ComUtilities.Release(ref connection);
            }
        });
    }

    /// <inheritdoc />
    public ConnectionAccountHintClearResult ClearAccountHint(IExcelBatch batch, string connectionName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(connectionName);
        ValidateAccountSettingsReadiness(batch);
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledb = null;
            try
            {
                connection = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName)
                    ?? throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                oledb = GetAccountSettingsConnection(connection);
                var settings = ReadAccountSettingsConnectionString(oledb);
                bool changed = HasAccountHint(settings);
                if (changed)
                {
                    settings.Remove("User ID");
                    settings.Remove("UID");
                    ct.ThrowIfCancellationRequested();
                    WriteAccountSettingsAndVerify(oledb, settings);
                }

                return new ConnectionAccountHintClearResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ConnectionName = connectionName,
                    Changed = changed,
                    AccountHintPresent = false
                };
            }
            finally
            {
                ComUtilities.Release(ref oledb);
                ComUtilities.Release(ref connection);
            }
        });
    }

    /// <summary>
    /// Checks native readiness before Service write-access checks can queue behind blocked Excel COM.
    /// </summary>
    public static void ValidateAccountSettingsReadiness(IExcelBatch batch)
    {
        var state = batch is IExcelBatchRefreshState refreshState
            ? refreshState.GetRefreshState()
            : WorkbookRefreshState.Unknown;
        ExcelBusyException.ThrowIfNotReady(state, "inspect or change connection account settings");
    }

    internal static void ValidateAccountSettingsUpdate(
        string? accountHint, ConnectionInteractiveLogin? interactiveLogin, ConnectionIdentityMode? identityMode)
    {
        if (accountHint == null && interactiveLogin == null && identityMode == null)
            throw new ArgumentException("Supply at least one accountHint, interactiveLogin or identityMode setting.");
        if (accountHint != null && (string.IsNullOrWhiteSpace(accountHint) || accountHint.Contains('\0')))
            throw new ArgumentException("accountHint must be nonblank and contain no null characters. Use clear-account-hint to remove it.");
        if (interactiveLogin.HasValue && !Enum.IsDefined(interactiveLogin.Value))
            throw new ArgumentException("interactiveLogin must be Default, Enabled, Disabled or Always.");
        if (identityMode.HasValue && !Enum.IsDefined(identityMode.Value))
            throw new ArgumentException("identityMode must be Default, CurrentUser, Connection or Process.");
    }

    private static void PopulateAccountSettingsResult(
        ConnectionAccountSettingsResult result, IExcelBatch batch, string name,
        Excel.OLEDBConnection oledb, DbConnectionStringBuilder settings)
    {
        result.FilePath = batch.WorkbookPath;
        result.ConnectionName = name;
        result.AccountHintPresent = HasAccountHint(settings);
        result.PasswordPresent = settings.ContainsKey("Password") || settings.ContainsKey("PWD");
        result.ImpersonationPresent = settings.ContainsKey("EffectiveUserName");
        result.SavePassword = oledb.SavePassword;
        result.InteractiveLogin = ReadSignInMode(settings, "Interactive Login", ["Default", "Enabled", "Disabled", "Always"]);
        result.IdentityMode = ReadSignInMode(settings, "Identity Mode", ["Default", "CurrentUser", "Connection", "Process"]);
        result.Success = true;
    }

    private static bool AccountSettingsMatch(DbConnectionStringBuilder actual, DbConnectionStringBuilder expected) =>
        actual.Count == expected.Count && expected.Keys.Cast<string>()
            .All(key => actual.ContainsKey(key) && Equals(expected[key], actual[key]));

    private static void WriteAccountSettingsAndVerify(Excel.OLEDBConnection oledb, DbConnectionStringBuilder expected)
    {
        oledb.Connection = "OLEDB;" + expected.ConnectionString;
        DbConnectionStringBuilder actual;
        try
        {
            actual = ParseAccountSettingsConnectionString(
                Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture) ?? "");
        }
        catch (Exception ex) when (ex is COMException or ArgumentException or NotSupportedException)
        {
            throw new InvalidOperationException(
                "Could not verify the connection after updating its account settings. " +
                "The workbook may have changed and has not been saved by this action. Inspect it before retrying.", ex);
        }
        if (!AccountSettingsMatch(actual, expected))
        {
            throw new InvalidOperationException(
                "Account-setting readback did not match the expected connection settings. " +
                "The workbook may have changed and has not been saved by this action. Inspect it before retrying.");
        }
    }

    private static Excel.OLEDBConnection GetAccountSettingsConnection(Excel.WorkbookConnection connection)
    {
        if (connection.Type != Excel.XlConnectionType.xlConnectionTypeOLEDB ||
            PowerQueryHelpers.IsPowerQueryConnection(connection))
        {
            throw new NotSupportedException("Account settings support only MSOLAP OLEDB connections, not Power Query or other connection types.");
        }
        return connection.OLEDBConnection;
    }

    private static DbConnectionStringBuilder ReadAccountSettingsConnectionString(Excel.OLEDBConnection oledb)
    {
        ValidateAccountSettingsConnectionReadiness(oledb.Refreshing);
        string text = Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture) ?? "";
        return ParseAccountSettingsConnectionString(text);
    }

    internal static void ValidateAccountSettingsConnectionReadiness(bool isRefreshing)
    {
        ExcelBusyException.ThrowIfNotReady(
            isRefreshing ? WorkbookRefreshState.Refreshing : WorkbookRefreshState.Ready,
            "inspect or change connection account settings");
    }

    internal static DbConnectionStringBuilder ParseAccountSettingsConnectionString(string text)
    {
        var settings = new DbConnectionStringBuilder
        {
            ConnectionString = text.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase) ? text[6..] : text
        };
        string provider = settings.TryGetValue("Provider", out var value)
            ? Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
            : "";
        if (!provider.Equals("MSOLAP", StringComparison.OrdinalIgnoreCase) &&
            !(provider.StartsWith("MSOLAP.", StringComparison.OrdinalIgnoreCase) &&
              provider[7..].Length > 0 && provider[7..].All(char.IsAsciiDigit)))
        {
            throw new NotSupportedException("Account settings support only MSOLAP OLEDB connections. No settings have been changed.");
        }
        return settings;
    }

    private static bool HasAccountHint(DbConnectionStringBuilder settings) =>
        settings.ContainsKey("User ID") || settings.ContainsKey("UID");

    internal static string? ReadSignInMode(DbConnectionStringBuilder settings, string key, string[] allowed)
    {
        if (!settings.TryGetValue(key, out var value)) return null;
        string? mode = Convert.ToString(value, CultureInfo.InvariantCulture);
        return allowed.FirstOrDefault(candidate => candidate.Equals(mode, StringComparison.OrdinalIgnoreCase)) ?? "Unrecognized";
    }
}
