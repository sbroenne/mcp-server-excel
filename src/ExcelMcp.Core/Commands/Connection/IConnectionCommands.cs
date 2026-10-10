using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data connections (OLEDB, ODBC, ODC import) and read-only external OLAP schema discovery.
/// TEXT/WEB/CSV: Use querytable for direct local imports or powerquery for transformations.
/// Power Query connections auto-redirect to powerquery.
/// TIMEOUT: Refresh accepts a caller timeout; load-to uses the 30-minute data-operation timeout.
/// </summary>
[ServiceCategory("Connection")]
[McpTool("connection", Title = "Data Connection Operations", Destructive = true, Category = "query",
    Description = "Create, change, import, delete, load, refresh, and cancel refreshes for data connections (OLEDB, ODBC, ODC import). Discover external OLAP dimensions, hierarchies, levels, and bounded member pages through an existing connection without refreshing data. Use querytable for direct text/web/CSV imports or powerquery for transformations. Power Query connections redirect by exact mashup Location identity. Delete/load-to cleanup follows the exact WorkbookConnection and preserves unrelated similarly named QueryTables. OLAP/MSOLAP connections refresh synchronously; background mode cannot be enabled. Refresh cancellation uses typed OLEDB/ODBC helpers. Refresh accepts a caller timeout; load-to uses the 30-minute data-operation timeout.")]
[McpReadOnlyActions("list", "view", "test", "get-refresh-status", "get-properties", "get-account-settings", "discover-olap-schema", "search-olap-members")]
public interface IConnectionCommands
{
    /// <summary>
    /// Lists all connections in a workbook
    /// </summary>
    [ServiceAction("list")]
    ConnectionListResult List(IExcelBatch batch);

    /// <summary>
    /// Discovers dimensions, hierarchies, and levels exposed by an existing external OLAP workbook connection.
    /// Uses the connection's current Excel authentication context and does not refresh data or modify the workbook.
    /// Embedded Data Model connections, Power Query connections, and non-OLAP providers are unsupported.
    /// Unique names are returned separately from provider display captions. Connection strings and credentials are never returned.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Exact name of the existing external OLAP workbook connection</param>
    /// <param name="hierarchyUniqueName">Optional exact provider hierarchy unique name filter</param>
    /// <param name="levelUniqueName">Optional exact provider level unique name filter; must belong to the selected hierarchy when one is supplied</param>
    [ServiceAction("discover-olap-schema")]
    ExternalOlapSchemaResult DiscoverOlapSchema(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        string? hierarchyUniqueName = null,
        string? levelUniqueName = null);

    /// <summary>
    /// Searches or lists a bounded page of members at an external OLAP hierarchy level.
    /// Uses the existing workbook connection's Excel authentication context without refreshing or changing PivotTable filters.
    /// Continuation tokens are opaque and bound to the connection, hierarchy, level, search text, and page size.
    /// Provider totals are returned when available; totals for text-filtered searches may be unknown.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Exact name of the existing external OLAP workbook connection</param>
    /// <param name="hierarchyUniqueName">Exact provider hierarchy unique name</param>
    /// <param name="levelUniqueName">Exact provider level unique name in that hierarchy</param>
    /// <param name="searchText">Optional case-insensitive text filter for member names, captions, and unique names</param>
    /// <param name="continuationToken">Opaque token returned by a previous page request for the same query</param>
    /// <param name="pageSize">Number of matching members to return, from 1 through 250; defaults to 100</param>
    [ServiceAction("search-olap-members")]
    ExternalOlapMemberSearchResult SearchOlapMembers(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [RequiredParameter] string hierarchyUniqueName,
        [RequiredParameter] string levelUniqueName,
        string? searchText = null,
        string? continuationToken = null,
        int pageSize = 100);

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to view</param>
    [ServiceAction("view")]
    ConnectionViewResult View(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name for the new connection</param>
    /// <param name="connectionString">OLEDB or ODBC connection string</param>
    /// <param name="commandText">SQL query or table name</param>
    /// <param name="description">Optional description for the connection</param>
    [ServiceAction("create")]
    OperationResult Create(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [RequiredParameter, FromString("connectionString")] string connectionString,
        [FromString("commandText")] string? commandText = null,
        [FromString("description")] string? description = null);

    /// <summary>
    /// Refreshes connection data with optional timeout
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to refresh</param>
    /// <param name="timeout">Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch</param>
    [ServiceAction("refresh")]
    OperationResult Refresh(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        TimeSpan? timeout = null);

    /// <summary>
    /// Gets refresh status for OLEDB and ODBC background refreshes started
    /// outside the synchronous connection refresh action.
    /// Excel PIA exposes status on the typed sub-connection, not WorkbookConnection.
    /// During a synchronous Excel operation, status returns Busy instead of waiting in the session queue.
    /// A failed status read is an error, not evidence of completion.
    /// </summary>
    [ServiceAction("get-refresh-status")]
    RefreshStatusResult GetRefreshStatus(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Cancels an active OLEDB or ODBC background refresh started outside the
    /// synchronous connection refresh action.
    /// Returns an explicit unsupported result for connection types without a typed PIA cancellation API.
    /// Returns Busy while a synchronous Excel operation prevents safe access to the session.
    /// </summary>
    [ServiceAction("cancel-refresh")]
    RefreshCancellationResult CancelRefresh(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Deletes a connection and QueryTables owned by that exact WorkbookConnection.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to delete</param>
    [ServiceAction("delete")]
    OperationResult Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet, replacing only QueryTables owned by
    /// the exact WorkbookConnection.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    /// <param name="sheetName">Target worksheet name</param>
    [ServiceAction("load-to")]
    OperationResult LoadTo(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [RequiredParameter, FromString("sheetName")] string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    [ServiceAction("get-properties")]
    ConnectionPropertiesResult GetProperties(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Inspects saved account-hint, password, impersonation and sign-in settings for an MSOLAP OLEDB connection.
    /// Returns presence flags and recognized sign-in modes, never account names, passwords or tokens.
    /// Power Query, ODBC and other providers are unsupported. Requires idle Excel.
    /// Does not sign in, refresh, save, or inspect Office/Windows credential caches.
    /// Unconfigured sign-in modes are null; unrecognized values are reported as Unrecognized without exposing them.
    /// </summary>
    [ServiceAction("get-account-settings")]
    ConnectionAccountSettingsResult GetAccountSettings(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Sets only explicitly supplied account-hint and sign-in settings on the selected MSOLAP OLEDB connection.
    /// Requires at least one setting and idle, writable Excel; omitted settings are preserved.
    /// Does not set passwords, tokens or EffectiveUserName impersonation; all unrelated properties are verified unchanged.
    /// Returns changed=false when the requested settings already match; account values are never returned.
    /// Does not sign in, sign out, clear shared credentials, force account selection, refresh or save.
    /// Explicit User ID overrides Identity Mode. Other connection settings can affect Interactive Login behavior.
    /// Power Query, ODBC and other providers are unsupported. A readback failure can leave changes in the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="connectionName">Exact workbook connection name.</param>
    /// <param name="accountHint">Nonblank User ID to store; null preserves the existing hint. Replaces User ID/UID aliases. Use clear-account-hint to remove a hint. Not echoed in results.</param>
    /// <param name="interactiveLogin">MSOLAP interactive sign-in mode: Default, Enabled, Disabled, Always. Null preserves the current setting.</param>
    /// <param name="identityMode">MSOLAP identity selection: Default, CurrentUser, Connection, Process. Null preserves the current setting. Explicit User ID takes precedence.</param>
    [ServiceAction("set-account-settings")]
    ConnectionAccountSettingsUpdateResult SetAccountSettings(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        string? accountHint = null,
        ConnectionInteractiveLogin? interactiveLogin = null,
        ConnectionIdentityMode? identityMode = null);

    /// <summary>
    /// Removes only User ID/UID account hints from the selected MSOLAP OLEDB connection and verifies readback.
    /// Preserves passwords, tokens, EffectiveUserName impersonation and all unrelated connection settings.
    /// Requires idle, writable Excel. Power Query, ODBC and other providers are unsupported.
    /// Does not sign out, clear Office/Windows credential caches, force account selection, refresh or save.
    /// Returns changed=false when no account hint exists. Save explicitly to persist the change.
    /// A readback failure can leave the workbook changed; inspect it rather than assuming rollback.
    /// </summary>
    [ServiceAction("clear-account-hint")]
    ConnectionAccountHintClearResult ClearAccountHint(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Sets connection properties (connection string, command text, description, and behavior settings)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    /// <param name="connectionString">New connection string (null to keep current)</param>
    /// <param name="commandText">New SQL query or table name (null to keep current)</param>
    /// <param name="description">New description (null to keep current)</param>
    /// <param name="backgroundQuery">Run query in background for non-OLAP connections (null to keep current). OLAP always refreshes synchronously; true is rejected before property changes, false skips the unsupported setting. Change an OLEDB provider and background mode in separate calls; combined provider transitions are rejected before any writes.</param>
    /// <param name="refreshOnFileOpen">Refresh when file opens (null to keep current)</param>
    /// <param name="savePassword">Save password in connection (null to keep current)</param>
    /// <param name="refreshPeriod">Nonnegative auto-refresh interval in minutes; 0 disables automatic refresh (null to keep current)</param>
    [ServiceAction("set-properties")]
    OperationResult SetProperties(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        string? connectionString = null,
        string? commandText = null,
        string? description = null,
        bool? backgroundQuery = null,
        bool? refreshOnFileOpen = null,
        bool? savePassword = null,
        int? refreshPeriod = null);

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to test</param>
    [ServiceAction("test")]
    OperationResult Test(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);
}
