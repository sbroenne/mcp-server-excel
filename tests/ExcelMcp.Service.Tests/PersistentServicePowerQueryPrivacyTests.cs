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
        Assert.Contains("ignored", enabledRead.Message, StringComparison.OrdinalIgnoreCase);

        await _fixture.SaveAndReopenAsync();

        var reopened = RequireSuccess(_queries.GetPrivacySettings(_fixture.BatchToken));
        Assert.True(reopened.IgnorePrivacyLevels);

        var disabled = RequireSuccess(
            _queries.SetPrivacySettings(_fixture.BatchToken, ignorePrivacyLevels: false));
        Assert.False(disabled.IgnorePrivacyLevels);
        Assert.Contains("not ignored", disabled.Message, StringComparison.OrdinalIgnoreCase);

        var disabledRead = RequireSuccess(_queries.GetPrivacySettings(_fixture.BatchToken));
        Assert.False(disabledRead.IgnorePrivacyLevels);
    }
}
