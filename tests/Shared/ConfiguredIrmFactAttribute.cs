using Xunit;

namespace Sbroenne.ExcelMcp.Tests.Helpers;

/// <summary>
/// Skips an opt-in IRM integration test unless TEST_IRM_FILE names an existing workbook.
/// </summary>
public sealed class ConfiguredIrmFactAttribute : FactAttribute
{
    private const string MissingFixtureMessage =
        "Set TEST_IRM_FILE to an existing IRM/AIP-protected workbook to run this opt-in regression.";

    /// <summary>
    /// Initializes a new instance of the <see cref="ConfiguredIrmFactAttribute"/> class.
    /// </summary>
    public ConfiguredIrmFactAttribute()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        if (string.IsNullOrWhiteSpace(irmTestFile) || !File.Exists(irmTestFile))
        {
            Skip = MissingFixtureMessage;
        }
    }
}
