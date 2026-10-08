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
                return new ConnectionAccountSettingsResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ConnectionName = connectionName,
                    AccountHintPresent = HasAccountHint(settings),
                    PasswordPresent = settings.ContainsKey("Password") || settings.ContainsKey("PWD"),
                    ImpersonationPresent = settings.ContainsKey("EffectiveUserName"),
                    SavePassword = oledb.SavePassword,
                    InteractiveLogin = ReadSignInMode(settings, "Interactive Login", ["Default", "Enabled", "Disabled", "Always"]),
                    IdentityMode = ReadSignInMode(settings, "Identity Mode", ["Default", "CurrentUser", "Connection", "Process"])
                };
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
                    oledb.Connection = "OLEDB;" + settings.ConnectionString;
                    DbConnectionStringBuilder actual;
                    try
                    {
                        actual = ParseAccountSettingsConnectionString(
                            Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture) ?? "");
                    }
                    catch (Exception ex) when (ex is COMException or ArgumentException or NotSupportedException)
                    {
                        throw new InvalidOperationException(
                            "Could not verify the connection after removing its account hint. " +
                            "The workbook may have changed and has not been saved by this action. Inspect it before retrying.", ex);
                    }
                    if (actual.Count != settings.Count || settings.Keys.Cast<string>()
                        .Any(key => !actual.ContainsKey(key) || !Equals(settings[key], actual[key])))
                    {
                        throw new InvalidOperationException(
                            "Account-hint readback did not match the expected connection settings. " +
                            "The workbook may have changed and has not been saved by this action. Inspect it before retrying.");
                    }
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
        if (oledb.Refreshing)
            throw new InvalidOperationException("The connection is refreshing. Wait until it is idle before inspecting or changing account settings.");
        string text = Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture) ?? "";
        return ParseAccountSettingsConnectionString(text);
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
