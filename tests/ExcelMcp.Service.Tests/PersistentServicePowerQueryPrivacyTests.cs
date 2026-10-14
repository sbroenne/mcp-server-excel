using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryPrivacyTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();

    [Fact]
    public async Task PrivacySettings_SetAndReadBothValues_PersistsWhenWorkbookIsSaved()
    {
        var enabled = RequireSuccess(
            _queries.SetPrivacySettings(_fixture.BatchToken, ignorePrivacyLevels: true));
        Assert.True(enabled.IgnorePrivacyLevels);
        Assert.Contains(
            "Data from different sources may be combined and sent between them.",
            enabled.Message,
            StringComparison.Ordinal);

        var enabledRead = RequireSuccess(_queries.GetPrivacySettings(_fixture.BatchToken));
        Assert.True(enabledRead.IgnorePrivacyLevels);
        Assert.Contains("configured to ignore", enabledRead.Message, StringComparison.Ordinal);
        Assert.Contains("global privacy option may override", enabledRead.Message, StringComparison.Ordinal);

        await _fixture.SaveAndReopenAsync();

        var reopened = RequireSuccess(_queries.GetPrivacySettings(_fixture.BatchToken));
        Assert.True(reopened.IgnorePrivacyLevels);

        var disabled = RequireSuccess(
            _queries.SetPrivacySettings(_fixture.BatchToken, ignorePrivacyLevels: false));
        Assert.False(disabled.IgnorePrivacyLevels);
        Assert.Contains("configured not to ignore", disabled.Message, StringComparison.Ordinal);
        Assert.Contains("global privacy option may override", disabled.Message, StringComparison.Ordinal);

        var disabledRead = RequireSuccess(_queries.GetPrivacySettings(_fixture.BatchToken));
        Assert.False(disabledRead.IgnorePrivacyLevels);

        await _fixture.SaveAndReopenAsync();

        var reopenedDisabled = RequireSuccess(_queries.GetPrivacySettings(_fixture.BatchToken));
        Assert.False(reopenedDisabled.IgnorePrivacyLevels);
    }
}
