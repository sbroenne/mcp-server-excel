using System.Text;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

public sealed class McpToolProfileTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("full")]
    [InlineData("unknown-profile")]
    public void Resolve_MissingFullOrUnknown_PreservesFullCompatibility(string? value)
    {
        Assert.Equal(McpToolProfile.Full, McpToolProfileCatalog.Resolve(value));
    }

    [Theory]
    [InlineData("compact")]
    [InlineData("copilot-compact")]
    [InlineData(" COPILOT-COMPACT ")]
    public void Resolve_CompactAliases_SelectCopilotCompact(string value)
    {
        Assert.Equal(McpToolProfile.CopilotCompact, McpToolProfileCatalog.Resolve(value));
    }

    [Fact]
    public void CopilotCompactProfile_IsSmallStableAndFormattingCapable()
    {
        Assert.Equal(
            [
                "file",
                "workflow",
                "worksheet",
                "range",
                "range_edit",
                "range_format",
                "worksheet_style",
                "layout",
                "calculation_mode",
            ],
            McpToolProfileCatalog.CompactTools);
        Assert.Equal(
            McpToolProfileCatalog.CompactTools.Count,
            McpToolProfileCatalog.CompactTools.Distinct(StringComparer.Ordinal).Count());
        Assert.True(Encoding.UTF8.GetByteCount(McpToolProfileCatalog.CompactServerInstructions) < 900);
        Assert.Contains("tools/list", McpToolProfileCatalog.CompactServerInstructions, StringComparison.Ordinal);
        Assert.Contains("unknown outcome", McpToolProfileCatalog.CompactServerInstructions, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("EXCELMCP_TOOL_PROFILE=full", McpToolProfileCatalog.CompactServerInstructions, StringComparison.Ordinal);
    }
}
