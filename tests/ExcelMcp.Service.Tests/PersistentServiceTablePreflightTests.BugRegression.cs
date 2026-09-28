using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void Append_WithJsonElementValues_DoesNotThrow()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        _rangeCommands.SetValues(
            batch,
            "Data",
            "A1:C2",
            [["Label", "IsActive", "Amount"], ["Initial", true, 1.0]]);
        _tableCommands.Create(
            batch, "Data", "DataTable", "A1:C2", true, "TableStyleLight1");

        var rowsJson =
            """[["NewRow", true, 99.5], ["AnotherRow", false, 0.0]]""";
        var deserializedRows =
            JsonSerializer.Deserialize<List<List<object?>>>(rowsJson)!;

        Assert.IsType<JsonElement>(deserializedRows[0][0]);
        Assert.IsType<JsonElement>(deserializedRows[0][1]);
        Assert.IsType<JsonElement>(deserializedRows[0][2]);

        _tableCommands.Append(batch, "DataTable", deserializedRows);

        var info = _tableCommands.Read(batch, "DataTable");
        Assert.True(info.Success, $"Read after append failed: {info.ErrorMessage}");
        Assert.Equal(3, info.Table!.RowCount);
    }
}
