using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacWorksheetLookupTests
{
    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("History")]
    [InlineData("'Name")]
    [InlineData("Name'")]
    [InlineData("Bad/Name")]
    [InlineData("12345678901234567890123456789012")]
    public void Create_InvalidNameFailsBeforeAnyExcelDispatch(string name)
    {
        var dispatched = 0;
        var backend = new MacExcelBackend((_, _, _) =>
        {
            dispatched++;
            return Task.FromResult(new MacProcessResult(0, """{"success":true}""", ""));
        });
        var session = new MacExcelSession
        {
            SessionId = "name-validation",
            FilePath = Path.GetFullPath("opaque.xlsx"),
            IsVisible = false,
            OperationTimeout = TimeSpan.FromSeconds(10)
        };
        using var batch = new MacExcelBatch(backend, session);
        var error = Assert.Throws<ArgumentException>(() =>
            batch.Invoke<OperationResult>("sheet.create", new { sheetName = name }));
        Assert.Equal("name", error.ParamName);
        Assert.Equal(0, dispatched);
    }

    [Theory]
    [InlineData("""{"success":true}""")]
    [InlineData("""{"success":true,"worksheets":null}""")]
    [InlineData("""{"success":true,"worksheets":[{}]}""")]
    public void MalformedDiscoveryIsNotReportedAsAnOrdinaryMissingSheet(string response)
    {
        var backend = new MacExcelBackend((_, _, _) => Task.FromResult(new MacProcessResult(0, response, "")));
        var session = new MacExcelSession
        {
            SessionId = "malformed-worksheet-lookup",
            FilePath = Path.GetFullPath("opaque.xlsx"),
            IsVisible = false,
            OperationTimeout = TimeSpan.FromSeconds(10)
        };
        using var batch = new MacExcelBatch(backend, session);
        Assert.Throws<InvalidDataException>(() =>
            batch.Invoke<OperationResult>("sheet.rename", new { oldName = "Sheet1", newName = "Changed" }));
    }

    [Theory]
    [InlineData("sheet.rename", "sheet1")]
    [InlineData("sheet.rename", "'Sheet1'")]
    [InlineData("sheet.rename", "Missing")]
    [InlineData("sheet.delete", "sheet1")]
    [InlineData("sheet.delete", "'Sheet1'")]
    [InlineData("sheet.delete", "Missing")]
    public void MissingExactNameFailsBeforeMutationWithCoreDiagnostic(string command, string name)
    {
        var mutations = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (!start.ArgumentList.Contains("sheet.list")) mutations++;
            return Task.FromResult(new MacProcessResult(0,
                """{"success":true,"worksheets":[{"name":"Sheet1","index":1,"visible":true}]}""", ""));
        });
        var session = new MacExcelSession
        {
            SessionId = "strict-worksheet-lookup",
            FilePath = Path.GetFullPath("opaque.xlsx"),
            IsVisible = false,
            OperationTimeout = TimeSpan.FromSeconds(10)
        };
        using var batch = new MacExcelBatch(backend, session);
        var arguments = command == "sheet.rename"
            ? (object)new { oldName = name, newName = "Changed" }
            : new { sheetName = name };
        var error = Assert.Throws<InvalidOperationException>(() => batch.Invoke<OperationResult>(command, arguments));
        Assert.Equal($"Sheet '{name}' not found.", error.Message);
        Assert.Equal(0, mutations);
    }
}
