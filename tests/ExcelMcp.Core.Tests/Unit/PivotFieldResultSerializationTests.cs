using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class PivotFieldResultSerializationTests
{
    [Fact]
    public void RenderedText_SerializesForServiceClients()
    {
        var json = ServiceProtocol.Serialize(new PivotFieldResult
        {
            Success = true,
            NumberFormat = "0.00",
            RenderedText = "49.000,00"
        });

        using var document = JsonDocument.Parse(json);
        var result = document.RootElement;
        Assert.Equal("49.000,00", result.GetProperty("renderedText").GetString());
    }

    [Fact]
    public void RenderedText_IsOmittedWhenNoDataCellIsAvailable()
    {
        var json = ServiceProtocol.Serialize(new PivotFieldResult
        {
            Success = true,
            NumberFormat = "0.00"
        });

        using var document = JsonDocument.Parse(json);
        Assert.False(document.RootElement.TryGetProperty("renderedText", out _));
    }
}
