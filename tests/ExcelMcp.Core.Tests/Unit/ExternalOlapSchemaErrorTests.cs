using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ExternalOlapSchemaErrorTests
{
    private const string Sentinel = "sentinel-provider-secret";

    [Fact]
    public void AdoSessionFailureExplainsHowToRetryWithoutRefreshingOrExposingCredentials()
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "GetAdoConnectionUnavailableMessage", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var message = Assert.IsType<string>(method.Invoke(null, ["ExternalOlap"]));

        Assert.Contains("Do not refresh automatically", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("user authorizes", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("connection_name", message, StringComparison.Ordinal);
        Assert.Contains("excelcli connection refresh --connection-name", message, StringComparison.Ordinal);
        Assert.Contains("check that refresh succeeded", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("may update workbook data", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("does not guarantee", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("does not refresh data", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("does not prove", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("provider does not support this discovery path", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ExternalOlap", message, StringComparison.Ordinal);
        Assert.DoesNotContain(Sentinel, message, StringComparison.Ordinal);
    }

    [Fact]
    public void SchemaRowset_MembersReadOnlyRequiredColumns()
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "OpenSchemaRowset", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var connection = new MemberSchemaConnection();

        var rows = Assert.IsType<List<IReadOnlyDictionary<string, object?>>>(method.Invoke(
            null, [connection, ExternalOlapSchemaMapper.BuildMembersRequest("Cube", "[Hierarchy]", "[Level]"),
                "Cube", CancellationToken.None, 10001, null]));

        var row = Assert.Single(rows);
        Assert.Equal(2, row.Count);
        Assert.Equal("[Member]", row["member_unique_name"]);
        Assert.Equal("Member caption", row["MEMBER_CAPTION"]);
        Assert.True(connection.Recordset.Closed);
    }

    [Fact]
    public void SchemaRowset_MemberReadsStopAtRequestedBound()
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "OpenSchemaRowset", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var connection = new MemberSchemaConnection(rowCount: 10002);

        var rows = Assert.IsType<List<IReadOnlyDictionary<string, object?>>>(method.Invoke(
            null, [connection, ExternalOlapSchemaMapper.BuildMembersRequest("Cube", "[Hierarchy]", "[Level]"),
                "Cube", CancellationToken.None, 10001, null]));

        Assert.Equal(10001, rows.Count);
        Assert.Equal(10001, connection.Recordset.RowsConsumed);
        Assert.False(connection.Recordset.EOF);
        Assert.True(connection.Recordset.Closed);
    }

    [Fact]
    public void SchemaRowset_CancellationPreventsProviderCall()
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "OpenSchemaRowset", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var connection = new MemberSchemaConnection();

        var invocation = Assert.Throws<TargetInvocationException>(() => method.Invoke(
            null, [connection, ExternalOlapSchemaMapper.BuildMembersRequest("Cube", "[Hierarchy]", "[Level]"),
                "Cube", new CancellationToken(canceled: true), 10001, null]));

        Assert.IsType<OperationCanceledException>(invocation.InnerException);
        Assert.False(connection.Opened);
    }

    public sealed class MemberSchemaConnection(int rowCount = 1)
    {
        public MemberSchemaRecordset Recordset { get; } = new(rowCount);
        public bool Opened { get; private set; }
        public object OpenSchema(int schema, object?[] restrictions)
        {
            Assert.Equal(38, schema);
            Assert.Equal("[Level]", restrictions[5]);
            Opened = true;
            return Recordset;
        }
    }

    public sealed class MemberSchemaRecordset(int rowCount)
    {
        public MemberSchemaFields Fields { get; } = new();
        public bool EOF => RowsConsumed == rowCount;
        public int RowsConsumed { get; private set; }
        public int State => Closed ? 0 : 1;
        public bool Closed { get; private set; }
        public void MoveNext() => RowsConsumed++;
        public void Close() => Closed = true;
    }

    public sealed class MemberSchemaFields
    {
        private readonly MemberSchemaField[] _fields =
        [
            new("member_unique_name", "[Member]"),
            new("MEMBER_CAPTION", "Member caption"),
            new("PROVIDER_SPECIFIC_PROPERTY", null)
        ];
        public int Count => _fields.Length;
        public MemberSchemaField Item(int index) => _fields[index];
    }

    public sealed class MemberSchemaField(string name, string? value)
    {
        public string Name => name;
        public object Value => value ?? throw new InvalidOperationException("Unused provider property must not be read.");
    }

    [Theory]
    [InlineData(unchecked((int)0x80070005), "denied access")]
    [InlineData(unchecked((int)0x800A0CB3), "does not support")]
    [InlineData(unchecked((int)0x80004005), "could not return OLAP metadata")]
    public void ProviderError_KeepsClassificationAndHresultButOmitsProviderText(int hresult, string classification)
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "OpenSchemaRowset", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var invocation = Assert.Throws<TargetInvocationException>(() => method.Invoke(
            null, [new ErrorSchemaConnection(hresult),
                ExternalOlapSchemaMapper.BuildDimensionRequest("Cube"), "Cube", CancellationToken.None, null, null]));
        var error = Assert.IsType<InvalidOperationException>(invocation.InnerException);

        Assert.Contains(classification, error.Message, StringComparison.Ordinal);
        Assert.Contains($"0x{hresult:X8}", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(Sentinel, error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private-host", error.Message, StringComparison.Ordinal);
        Assert.Equal($"0x{hresult:X8}", OperationFailureClassifier.GetComHResult(error));
        Assert.Equal("ComInterop", OperationFailureClassifier.Classify(error));
        Assert.DoesNotContain("private-host", error.ToString(), StringComparison.Ordinal);
        Assert.DoesNotContain(Sentinel, error.ToString(), StringComparison.Ordinal);
    }

    public sealed class ErrorSchemaConnection(int hresult)
    {
#pragma warning disable CA2201 // Simulated provider error at the schema COM-call boundary.
        public object OpenSchema(int schema, object?[] restrictions) =>
            throw new COMException($"Provider failure: Password={Sentinel};Data Source=private-host", hresult);
#pragma warning restore CA2201
    }

    [Theory]
    [InlineData(1)]
    [InlineData(9)]
    public void SchemaRowset_CloseFailureIsNotReportedAsSuccess(int state)
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "OpenSchemaRowset", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var invocation = Assert.Throws<TargetInvocationException>(() => method.Invoke(
            null, [new CloseErrorSchemaConnection(state: state), ExternalOlapSchemaMapper.BuildDimensionRequest("Cube"),
                "Cube", CancellationToken.None, null, null]));
        var error = Assert.IsType<InvalidOperationException>(invocation.InnerException);

        Assert.Contains("0x80004005", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(Sentinel, error.Message, StringComparison.Ordinal);
        Assert.Equal("0x80004005", OperationFailureClassifier.GetComHResult(error));
        Assert.Equal("ComInterop", OperationFailureClassifier.Classify(error));
        Assert.DoesNotContain(Sentinel, error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void SchemaRowset_CleanupFailurePreservesQueryCancellation()
    {
        var method = typeof(ConnectionCommands).GetMethod(
            "OpenSchemaRowset", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var invocation = Assert.Throws<TargetInvocationException>(() => method.Invoke(
            null, [new CloseErrorSchemaConnection(cancelRead: true),
                ExternalOlapSchemaMapper.BuildDimensionRequest("Cube"),
                "Cube", CancellationToken.None, null, null]));

        var canceled = Assert.IsType<OperationCanceledException>(invocation.InnerException);
        Assert.Equal("Query canceled.", canceled.Message);
        var failures = Assert.IsType<AggregateException>(canceled.InnerException);
        Assert.Equal(2, failures.InnerExceptions.Count);
        Assert.Equal("Query canceled.", Assert.IsType<OperationCanceledException>(failures.InnerExceptions[0]).Message);
        Assert.Equal("0x80004005", OperationFailureClassifier.GetComHResult(failures.InnerExceptions[1]));
        Assert.DoesNotContain(Sentinel, canceled.ToString(), StringComparison.Ordinal);
    }

    public sealed class CloseErrorSchemaConnection(bool cancelRead = false, int state = 1)
    {
        public EmptySchemaRecordset Recordset { get; } = new(unchecked((int)0x80004005), cancelRead, state);

        public object OpenSchema(int schema, object?[] restrictions)
        {
            Assert.Equal(33, schema);
            Assert.Equal("Cube", restrictions[2]);
            return Recordset;
        }
    }

    public sealed class EmptySchemaRecordset(int hresult, bool cancelRead, int state)
    {
        public EmptySchemaFields Fields { get; } = new();
        public bool EOF => cancelRead ? throw new OperationCanceledException("Query canceled.") : true;
        public int State { get; } = state;

#pragma warning disable CA2201 // Simulated provider cleanup error at the schema COM-call boundary.
        public void Close() => throw new COMException(Sentinel, hresult);
#pragma warning restore CA2201
    }

    public sealed class EmptySchemaFields
    {
        public int Count { get; }
    }
}
