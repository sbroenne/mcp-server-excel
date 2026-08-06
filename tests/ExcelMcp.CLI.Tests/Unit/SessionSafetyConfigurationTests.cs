using Sbroenne.ExcelMcp.Service.Safety;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "SafetyConfiguration")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SessionSafetyConfigurationTests
{
    [Theory]
    [InlineData("review")]
    [InlineData("checkpoint")]
    [InlineData("journal")]
    [InlineData("verification")]
    public void NormalizeAbnormalShutdownPolicy_EnablesRecoveryDiscardWhenSafetyControlIsEnabled(string control)
    {
        var configuration = control switch
        {
            "review" => new SessionSafetyConfiguration { ReviewMode = ReviewMode.Required },
            "checkpoint" => new SessionSafetyConfiguration { CheckpointMode = CheckpointMode.OnRequest },
            "journal" => new SessionSafetyConfiguration { JournalMode = JournalMode.On },
            "verification" => new SessionSafetyConfiguration { VerificationMode = VerificationMode.On },
            _ => throw new ArgumentOutOfRangeException(nameof(control), control, null)
        };

        var normalized = configuration.NormalizeAbnormalShutdownPolicy(abnormalShutdownPolicySpecified: false);

        Assert.Equal(AbnormalShutdownPolicy.DiscardWithRecoveryEvidence, normalized.AbnormalShutdownPolicy);
    }

    [Fact]
    public void NormalizeAbnormalShutdownPolicy_PreservesExplicitLegacyPolicy()
    {
        var configuration = new SessionSafetyConfiguration
        {
            ReviewMode = ReviewMode.Required,
            AbnormalShutdownPolicy = AbnormalShutdownPolicy.LegacyAutoSave
        };

        var normalized = configuration.NormalizeAbnormalShutdownPolicy(abnormalShutdownPolicySpecified: true);

        Assert.Equal(AbnormalShutdownPolicy.LegacyAutoSave, normalized.AbnormalShutdownPolicy);
    }

    [Fact]
    public void NormalizeAbnormalShutdownPolicy_PreservesLegacyDefaultWhenSafetyControlsAreOff()
    {
        var normalized = SessionSafetyConfiguration.Default.NormalizeAbnormalShutdownPolicy(abnormalShutdownPolicySpecified: false);

        Assert.Equal(AbnormalShutdownPolicy.LegacyAutoSave, normalized.AbnormalShutdownPolicy);
    }
}
