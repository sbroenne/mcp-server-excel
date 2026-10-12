using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Extensions.Logging;
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
    private const string AdoSessionPrerequisites =
        "To retry, authenticate or open the existing connection in Excel, then run discovery again. " +
        "Discovery does not refresh data. If Excel still does not expose an ADO session, this provider " +
        "does not support this discovery path; use provider-supported metadata access outside ExcelMcp. " +
        "Do not include connection strings or credentials in a request.";

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
                    ct, logger: batch.Logger);
                List<IReadOnlyDictionary<string, object?>> hierarchyRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildHierarchyRequest(cubeName),
                    connectionName,
                    ct, logger: batch.Logger);
                List<IReadOnlyDictionary<string, object?>> levelRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildLevelRequest(cubeName),
                    connectionName,
                    ct, logger: batch.Logger);

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
                    ct, logger: batch.Logger);
                List<IReadOnlyDictionary<string, object?>> levelRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildLevelRequest(cubeName, levelUniqueName),
                    connectionName,
                    ct, logger: batch.Logger);
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

                List<IReadOnlyDictionary<string, object?>> memberRows = OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildMembersRequest(
                        cubeName,
                        hierarchyUniqueName,
                        levelUniqueName),
                    connectionName,
                    ct,
                    ExternalOlapSchemaMapper.MaximumMemberRows + 1,
                    batch.Logger);
                List<ExternalOlapMemberInfo> mappedMembers = ExternalOlapSchemaMapper.MapMembers(memberRows);
                List<ExternalOlapMemberInfo> orderedMembers = ExternalOlapSchemaMapper.OrderMembersForPaging(mappedMembers, position);
                string memberListDigest = ExternalOlapSchemaMapper.GetMemberListDigest(orderedMembers);
                OlapMemberPage page = ExternalOlapSchemaMapper.SelectMemberPage(
                    orderedMembers.Skip(position?.Offset ?? 0).ToArray(),
                    pageSize,
                    normalizedSearch,
                    MaximumOlapMemberRowsScanned);

                result.Members = page.Members;
                result.ReturnedCount = page.Members.Count;
                result.ScannedCount = memberRows.Count;
                result.TotalCount = totalCount;
                result.OmittedCount = totalCount.HasValue
                    ? Math.Max(0, totalCount.Value - page.Members.Count)
                    : null;
                result.ContinuationToken = page.HasMore && page.LastConsumedUniqueName is not null
                    ? ExternalOlapSchemaMapper.CreateContinuationToken(
                        scope,
                        new OlapMemberPosition(
                            (position?.Offset ?? 0) + page.ScannedCount,
                            page.LastConsumedUniqueName,
                            memberListDigest))
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
                throw new InvalidOperationException(GetAdoConnectionUnavailableMessage(connectionName));
            }

            return adoConnection;
        }
        catch (COMException ex)
        {
#pragma warning disable CA2201 // Preserve the actual COM HRESULT without retaining sensitive provider text.
            throw new InvalidOperationException(
                $"{GetAdoConnectionUnavailableMessage(connectionName)} (HRESULT 0x{ex.HResult:X8}).",
                new COMException("Excel could not expose the OLAP ADO session.", ex.HResult));
#pragma warning restore CA2201
        }
        catch (RuntimeBinderException)
        {
            throw new NotSupportedException(
                GetAdoConnectionUnavailableMessage(connectionName));
        }
    }

    private static string GetAdoConnectionUnavailableMessage(string connectionName) =>
        $"Excel did not expose an active ADO session for OLAP connection '{connectionName}'. " +
        AdoSessionPrerequisites;

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

    // ADO schema rowsets have no referenced PIA; late binding works with installed OLAP providers.
    private static List<IReadOnlyDictionary<string, object?>> OpenSchemaRowset(
        dynamic adoConnection,
        OlapSchemaRequest request,
        string connectionName,
        CancellationToken cancellationToken,
        int? maximumRows = null,
        ILogger? logger = null)
    {
        dynamic? recordset = null;
        dynamic? fields = null;
        bool readCompleted = false;
        Exception? cleanupFailure = null;
        var rows = new List<IReadOnlyDictionary<string, object?>>();
        try
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                recordset = adoConnection.OpenSchema(request.Schema, request.Restrictions);
                if (recordset is null)
                {
                    throw new InvalidOperationException(
                        "The OLAP provider returned no result for a schema rowset request.");
                }

                fields = recordset.Fields;
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

                while ((maximumRows is null || rows.Count < maximumRows)
                    && !Convert.ToBoolean(recordset.EOF, CultureInfo.InvariantCulture))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < fieldCount; i++)
                    {
                        if (request.Schema == ExternalOlapSchemaMapper.AdSchemaMembers
                            && columnNames[i].ToUpperInvariant() is not ("MEMBER_UNIQUE_NAME" or "MEMBER_CAPTION"
                                or "MEMBER_NAME" or "MEMBER_ORDINAL" or "MEMBER_TYPE"))
                        {
                            continue;
                        }

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

                readCompleted = true;
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
            try
            {
                if (recordset is not null)
                {
                    try
                    {
                        if ((Convert.ToInt32(recordset.State, CultureInfo.InvariantCulture) & 1) != 0)
                        {
                            recordset.Close();
                        }
                    }
                    catch (COMException ex)
                    {
                        if (readCompleted)
                        {
                            cleanupFailure = CreateOlapQueryError(connectionName, ex.HResult, ex.Message);
                        }
                        else
                        {
                            logger?.LogWarning("OLAP schema rowset cleanup failed (HRESULT 0x{HResult:X8}); preserving the query failure.",
                                ex.HResult);
                        }
                    }
                    catch (RuntimeBinderException)
                    {
                        if (readCompleted)
                        {
                            cleanupFailure = new NotSupportedException("The OLAP provider could not close its schema rowset.");
                        }
                        else
                        {
                            logger?.LogWarning("The OLAP provider could not close its schema rowset; preserving the query failure.");
                        }
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref recordset);
            }
        }

        if (cleanupFailure is not null)
        {
            throw cleanupFailure;
        }

        return rows;
    }

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
#pragma warning disable CA2201 // Preserve the actual COM HRESULT without retaining sensitive provider text.
        return new InvalidOperationException(
            $"OLAP schema discovery failed for connection '{connectionName}'. {reason} "
            + $"(HRESULT 0x{hresult:X8}).",
            new COMException(reason, hresult));
#pragma warning restore CA2201
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
