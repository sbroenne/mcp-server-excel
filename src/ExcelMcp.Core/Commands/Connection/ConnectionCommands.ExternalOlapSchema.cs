using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class ConnectionCommands
{
    private const int MaximumOlapMemberPageSize = 250;
    private const int MaximumOlapMemberRowsScanned = 1000;

    /// <inheritdoc />
    public ExternalOlapSchemaResult DiscoverOlapSchema(
        IExcelBatch batch,
        string connectionName,
        string? hierarchyUniqueName = null,
        string? levelUniqueName = null)
    {
        ValidateConnectionName(connectionName);
        ValidateOptionalName(hierarchyUniqueName, nameof(hierarchyUniqueName));
        ValidateOptionalName(levelUniqueName, nameof(levelUniqueName));

        var result = new ExternalOlapSchemaResult
        {
            FilePath = batch.WorkbookPath,
            ConnectionName = connectionName
        };
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledbConnection = null;
            dynamic? adoConnection = null;
            try
            {
                connection = GetExternalOlapConnection(ctx.Book, connectionName, out oledbConnection);
                adoConnection = GetAdoConnection(connectionName, oledbConnection);
                string cubeName = GetExternalOlapCubeName(connectionName, oledbConnection);

                List<IReadOnlyDictionary<string, object?>> dimensionRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildDimensionRequest(cubeName),
                    connectionName,
                    ct);
                List<IReadOnlyDictionary<string, object?>> hierarchyRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildHierarchyRequest(cubeName),
                    connectionName,
                    ct);
                List<IReadOnlyDictionary<string, object?>> levelRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildLevelRequest(cubeName),
                    connectionName,
                    ct);

                ExternalOlapSchemaResult mapped = ExternalOlapSchemaMapper.MapSchema(
                    connectionName,
                    dimensionRows,
                    hierarchyRows,
                    levelRows,
                    hierarchyUniqueName,
                    levelUniqueName);
                mapped.FilePath = result.FilePath;
                mapped.Success = true;
                return mapped;
            }
            finally
            {
                ComUtilities.Release(ref adoConnection);
                ComUtilities.Release(ref oledbConnection);
                ComUtilities.Release(ref connection);
            }
        }, timeoutCts.Token);
    }

    /// <inheritdoc />
    public ExternalOlapMemberSearchResult SearchOlapMembers(
        IExcelBatch batch,
        string connectionName,
        string hierarchyUniqueName,
        string levelUniqueName,
        string? searchText = null,
        string? continuationToken = null,
        int pageSize = 100)
    {
        ValidateConnectionName(connectionName);
        ArgumentException.ThrowIfNullOrWhiteSpace(hierarchyUniqueName);
        ArgumentException.ThrowIfNullOrWhiteSpace(levelUniqueName);
        if (pageSize is < 1 or > MaximumOlapMemberPageSize)
        {
            throw new ArgumentOutOfRangeException(
                nameof(pageSize),
                pageSize,
                $"pageSize must be between 1 and {MaximumOlapMemberPageSize}.");
        }

        var normalizedSearch = string.IsNullOrWhiteSpace(searchText) ? null : searchText.Trim();
        var result = new ExternalOlapMemberSearchResult
        {
            FilePath = batch.WorkbookPath,
            ConnectionName = connectionName,
            HierarchyUniqueName = hierarchyUniqueName,
            LevelUniqueName = levelUniqueName,
            SearchText = normalizedSearch
        };
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledbConnection = null;
            dynamic? adoConnection = null;
            try
            {
                connection = GetExternalOlapConnection(ctx.Book, connectionName, out oledbConnection);
                adoConnection = GetAdoConnection(connectionName, oledbConnection);
                string cubeName = GetExternalOlapCubeName(connectionName, oledbConnection);
                var scope = new OlapMemberSearchScope(
                    connectionName,
                    cubeName,
                    hierarchyUniqueName,
                    levelUniqueName,
                    normalizedSearch,
                    pageSize);
                var position = continuationToken is null
                    ? null
                    : ExternalOlapSchemaMapper.ReadContinuationToken(continuationToken, scope);

                List<IReadOnlyDictionary<string, object?>> hierarchyRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildHierarchyRequest(cubeName, hierarchyUniqueName),
                    connectionName,
                    ct);
                List<IReadOnlyDictionary<string, object?>> levelRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildLevelRequest(cubeName, levelUniqueName),
                    connectionName,
                    ct);
                ExternalOlapSchemaResult levelSchema = ExternalOlapSchemaMapper.MapSchema(
                    connectionName,
                    Array.Empty<IReadOnlyDictionary<string, object?>>(),
                    hierarchyRows,
                    levelRows,
                    hierarchyUniqueName,
                    levelUniqueName);
                long? totalCount = normalizedSearch is null
                    ? levelSchema.Levels.Single().MemberCount
                    : null;

                int readLimit = normalizedSearch is null
                    ? pageSize + 1
                    : MaximumOlapMemberRowsScanned + 1;
                List<IReadOnlyDictionary<string, object?>> memberRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildMembersRequest(
                        cubeName,
                        hierarchyUniqueName,
                        levelUniqueName),
                    connectionName,
                    ct,
                    position,
                    readLimit);
                List<ExternalOlapMemberInfo> mappedMembers = ExternalOlapSchemaMapper.MapMembers(memberRows);
                OlapMemberPage page = ExternalOlapSchemaMapper.SelectMemberPage(
                    mappedMembers,
                    pageSize,
                    normalizedSearch,
                    MaximumOlapMemberRowsScanned);

                result.Members = page.Members;
                result.ReturnedCount = page.Members.Count;
                result.ScannedCount = page.ScannedCount;
                result.TotalCount = totalCount;
                result.OmittedCount = totalCount.HasValue
                    ? Math.Max(0, totalCount.Value - page.Members.Count)
                    : null;
                result.ContinuationToken = page.HasMore && page.LastConsumedUniqueName is not null
                    ? ExternalOlapSchemaMapper.CreateContinuationToken(
                        scope,
                        new OlapMemberPosition(
                            (position?.Offset ?? 0) + page.ScannedCount,
                            page.LastConsumedUniqueName))
                    : null;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref adoConnection);
                ComUtilities.Release(ref oledbConnection);
                ComUtilities.Release(ref connection);
            }
        }, timeoutCts.Token);
    }

    private static Excel.WorkbookConnection GetExternalOlapConnection(
        Excel.Workbook workbook,
        string connectionName,
        out Excel.OLEDBConnection? oledbConnection)
    {
        oledbConnection = null;
        var connection = PowerQueryHelpers.FindConnectionByExactName(workbook, connectionName);
        if (connection is null)
        {
            throw new InvalidOperationException($"Connection '{connectionName}' was not found.");
        }

        bool keepConnection = false;
        try
        {
            if (connection.Type != Excel.XlConnectionType.xlConnectionTypeOLEDB
                || PowerQueryHelpers.IsPowerQueryConnection(connection))
            {
                throw new InvalidOperationException(
                    $"Connection '{connectionName}' is unsupported; select an external OLAP OLE DB connection.");
            }

            oledbConnection = connection.OLEDBConnection;
            if (oledbConnection is null || !oledbConnection.OLAP)
            {
                throw new InvalidOperationException(
                    $"Connection '{connectionName}' is unsupported; select an external OLAP OLE DB connection.");
            }

            keepConnection = true;
            return connection;
        }
        finally
        {
            if (!keepConnection)
            {
                ComUtilities.Release(ref oledbConnection);
                ComUtilities.Release(ref connection);
            }
        }
    }

    private static dynamic GetAdoConnection(
        string connectionName,
        Excel.OLEDBConnection? oledbConnection)
    {
        try
        {
            dynamic? adoConnection = oledbConnection?.ADOConnection;
            if (adoConnection is null)
            {
                throw new InvalidOperationException(
                    $"Connection '{connectionName}' has no active Excel ADO session for OLAP schema discovery.");
            }

            return adoConnection;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                $"Excel did not expose an active ADO session for OLAP connection '{connectionName}'. "
                + $"Schema discovery requires the existing connection session (HRESULT 0x{ex.HResult:X8}).");
        }
        catch (RuntimeBinderException)
        {
            throw new NotSupportedException(
                $"Connection '{connectionName}' does not expose an Excel ADO session for OLAP schema discovery.");
        }
    }

    private static string GetExternalOlapCubeName(
        string connectionName,
        Excel.OLEDBConnection? oledbConnection)
    {
        if (oledbConnection is null || oledbConnection.CommandType != Excel.XlCmdType.xlCmdCube)
        {
            throw new InvalidOperationException(
                $"Connection '{connectionName}' does not select an OLAP cube command.");
        }

        string? cubeName = oledbConnection.CommandText switch
        {
            string text => text.Trim(),
            string[] { Length: 1 } text => text[0]?.Trim(),
            object[] { Length: 1 } values =>
                Convert.ToString(values[0], CultureInfo.InvariantCulture)?.Trim(),
            _ => null
        };
        if (string.IsNullOrWhiteSpace(cubeName))
        {
            throw new InvalidOperationException(
                $"Connection '{connectionName}' does not expose a selected OLAP cube name.");
        }

        return cubeName;
    }

    private static List<IReadOnlyDictionary<string, object?>> OpenSchemaRowset(
        dynamic adoConnection,
        OlapSchemaRequest request,
        string connectionName,
        CancellationToken cancellationToken,
        OlapMemberPosition? startAfter = null,
        int? maximumRows = null)
    {
        dynamic? recordset = null;
        dynamic? fields = null;
        try
        {
            try
            {
                recordset = adoConnection.OpenSchema(request.Schema, request.Restrictions);
                if (recordset is null)
                {
                    throw new InvalidOperationException(
                        "The OLAP provider returned no result for a schema rowset request.");
                }

                fields = recordset.Fields;
                if (startAfter is not null)
                {
                    SkipToPosition(recordset, fields, startAfter, cancellationToken);
                }

                int fieldCount = Convert.ToInt32(fields.Count, CultureInfo.InvariantCulture);
                var columnNames = new string[fieldCount];
                for (int i = 0; i < fieldCount; i++)
                {
                    dynamic? field = null;
                    try
                    {
                        field = fields.Item(i);
                        columnNames[i] = Convert.ToString(field.Name, CultureInfo.InvariantCulture)
                            ?? $"Column{i.ToString(CultureInfo.InvariantCulture)}";
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }

                var rows = new List<IReadOnlyDictionary<string, object?>>();
                while ((maximumRows is null || rows.Count < maximumRows)
                    && !Convert.ToBoolean(recordset.EOF, CultureInfo.InvariantCulture))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < fieldCount; i++)
                    {
                        dynamic? field = null;
                        try
                        {
                            field = fields.Item(i);
                            object? value = field.Value;
                            row[columnNames[i]] = value is DBNull ? null : value;
                        }
                        finally
                        {
                            ComUtilities.Release(ref field);
                        }
                    }

                    rows.Add(row);
                    recordset.MoveNext();
                }

                return rows;
            }
            catch (COMException ex)
            {
                throw CreateOlapQueryError(connectionName, ex.HResult, ex.Message);
            }
            catch (RuntimeBinderException)
            {
                throw new NotSupportedException(
                    $"The OLAP provider for connection '{connectionName}' does not expose the required schema rowset.");
            }
        }
        finally
        {
            if (recordset is not null)
            {
                try
                {
                    if (Convert.ToInt32(recordset.State, CultureInfo.InvariantCulture) == 1)
                    {
                        recordset.Close();
                    }
                }
                catch (Exception ex) when (ex is COMException or RuntimeBinderException)
                {
                    // Preserve the query result or primary provider failure; closing is cleanup only.
                }
            }

            ComUtilities.Release(ref fields);
            ComUtilities.Release(ref recordset);
        }
    }

    private static void SkipToPosition(
        dynamic recordset,
        dynamic fields,
        OlapMemberPosition position,
        CancellationToken cancellationToken)
    {
        for (int index = 0; index < position.Offset; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (Convert.ToBoolean(recordset.EOF, CultureInfo.InvariantCulture))
            {
                throw MemberListChanged();
            }

            if (index == position.Offset - 1)
            {
                dynamic? field = null;
                try
                {
                    field = fields.Item("MEMBER_UNIQUE_NAME");
                    string? uniqueName = Convert.ToString(field.Value, CultureInfo.InvariantCulture);
                    if (!string.Equals(uniqueName, position.LastUniqueName, StringComparison.Ordinal))
                    {
                        throw MemberListChanged();
                    }
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }

            recordset.MoveNext();
        }
    }

    private static InvalidOperationException MemberListChanged() =>
        new("The OLAP member list changed since the previous page was read. "
            + "Search again without continuationToken.");

    private static InvalidOperationException CreateOlapQueryError(
        string connectionName,
        int hresult,
        string? providerMessage)
    {
        string lowerMessage = providerMessage?.ToLowerInvariant() ?? string.Empty;
        string reason = hresult == unchecked((int)0x80070005)
            || hresult == unchecked((int)0x80040E4D)
            || hresult == unchecked((int)0x80040E09)
            || lowerMessage.Contains("permission", StringComparison.Ordinal)
            || lowerMessage.Contains("access denied", StringComparison.Ordinal)
            || lowerMessage.Contains("not authorized", StringComparison.Ordinal)
            || lowerMessage.Contains("login failed", StringComparison.Ordinal)
            ? "The current Excel user was denied access to OLAP metadata."
            : hresult == unchecked((int)0x800A0CB3)
                || lowerMessage.Contains("not recognized", StringComparison.Ordinal)
                || lowerMessage.Contains("not supported", StringComparison.Ordinal)
                || lowerMessage.Contains("unsupported", StringComparison.Ordinal)
                ? "The selected provider does not support the required OLAP schema rowset."
                : "The provider could not return OLAP metadata; verify schema-rowset support and the current user's metadata permissions.";
        return new InvalidOperationException(
            $"OLAP schema discovery failed for connection '{connectionName}'. {reason} "
            + $"(HRESULT 0x{hresult:X8}).");
    }

    private static void ValidateConnectionName(string connectionName)
    {
        if (string.IsNullOrWhiteSpace(connectionName))
        {
            throw new ArgumentException("connectionName is required.", nameof(connectionName));
        }
    }

    private static void ValidateOptionalName(string? value, string parameterName)
    {
        if (value is not null && string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException($"{parameterName} cannot be empty when supplied.", parameterName);
        }
    }
}
