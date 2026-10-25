using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceDocumentPropertyMetadataTests(
    PersistentServiceWorkbookFixture fixture,
    ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void DefaultList_FreshWorkbook_ReportsUnreadableBuiltInValueAndRetainsFilledControls()
    {
        const string title = "Native metadata control";
        const string customName = "MetadataControl";
        var commands = _fixture.CreateCommands<IWorkbookCommands>();
        RequireSuccess(commands.SetDocumentProperty(
            _fixture.BatchToken, "Title", title, DocumentPropertyScope.BuiltIn));
        RequireSuccess(commands.SetDocumentProperty(_fixture.BatchToken, customName, "custom control"));

        var native = _fixture.ExecuteRawVerification((context, _) =>
        {
            ExcelDocumentProperties? properties = null;
            ExcelDocumentProperties.Property? printDate = null;
            ExcelDocumentProperties.Property? titleProperty = null;
            try
            {
                properties = ExcelDocumentProperties.Open(context.Book, true);
                printDate = properties.Item("Last Print Date");
                titleProperty = properties.Item("Title");
                var failure = Assert.ThrowsAny<COMException>(() => { object? value = printDate.Value; });
                Assert.Equal(title, Assert.IsType<string>(titleProperty.Value));
                return (
                    Count: properties.Count,
                    Name: printDate.Name,
                    Type: (int)printDate.Type,
                    Linked: printDate.LinkToContent,
                    failure.HResult,
                    failure.Message);
            }
            finally
            {
                titleProperty?.Dispose();
                printDate?.Dispose();
                properties?.Dispose();
            }
        });
        output.WriteLine(
            $"Existing built-in '{native.Name}': Type={native.Type}, LinkToContent={native.Linked}; " +
            $"Value getter HRESULT=0x{native.HResult:X8}; {native.Message}");
        Assert.Equal("Last Print Date", native.Name, ignoreCase: true);
        Assert.Equal(3, native.Type);
        Assert.False(native.Linked);
        Assert.Equal(unchecked((int)0x80004005), native.HResult);

        // Omit both inclusion flags to exercise the advertised Service/MCP/CLI defaults.
        var response = _fixture.Send("workbook.list-document-properties", new { });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage), response.ErrorMessage);
        using var document = JsonDocument.Parse(response.Result!);
        var properties = document.RootElement.GetProperty("properties").EnumerateArray().ToArray();
        Assert.Equal(native.Count + 1, properties.Length);
        var print = Assert.Single(properties, property => property.GetProperty("name").GetString() == native.Name);
        AssertUnavailable(print, native.HResult, native.Message);
        Assert.Equal("date", print.GetProperty("valueType").GetString());
        Assert.Equal("built-in", print.GetProperty("scope").GetString());
        var filled = Assert.Single(properties, property => property.GetProperty("name").GetString() == "Title");
        AssertAvailable(filled, title, "built-in");
        var custom = Assert.Single(properties, property => property.GetProperty("name").GetString() == customName);
        AssertAvailable(custom, "custom control", "custom");

        var single = _fixture.Send("workbook.get-document-property", new
        {
            propertyName = native.Name,
            scope = "built-in"
        });
        Assert.True(single.Success, single.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(single.ErrorMessage), single.ErrorMessage);
        using var singleDocument = JsonDocument.Parse(single.Result!);
        AssertUnavailable(singleDocument.RootElement.GetProperty("property"), native.HResult, native.Message);
    }

    private static void AssertAvailable(JsonElement property, string value, string scope)
    {
        Assert.Equal(value, property.GetProperty("value").GetString());
        Assert.Equal("string", property.GetProperty("valueType").GetString());
        Assert.Equal(scope, property.GetProperty("scope").GetString());
        Assert.True(property.GetProperty("valueAvailable").GetBoolean());
        Assert.True(!property.TryGetProperty("valueError", out var error) || error.ValueKind == JsonValueKind.Null);
    }

    private static void AssertUnavailable(JsonElement property, int hresult, string message)
    {
        Assert.False(property.GetProperty("valueAvailable").GetBoolean());
        Assert.True(!property.TryGetProperty("value", out var value) || value.ValueKind == JsonValueKind.Null);
        var error = property.GetProperty("valueError").GetString();
        Assert.Contains($"0x{hresult:X8}", error, StringComparison.Ordinal);
        Assert.Contains(message, error, StringComparison.Ordinal);
        Assert.Contains("undefined", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("failed", error, StringComparison.OrdinalIgnoreCase);
    }
}
