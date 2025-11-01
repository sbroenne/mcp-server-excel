using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.VbaTrust;

/// <summary>
/// Tests for VbaTrustRequiredResult model and TestVbaTrustScope helper
/// </summary>
public partial class VbaTrustDetectionTests
{
    [Fact]
    public void VbaTrustRequiredResult_HasAllRequiredProperties()
    {
        // Act
        var trustResult = new VbaTrustRequiredResult
        {
            Success = false,
            ErrorMessage = "VBA trust access is not enabled",
            IsTrustEnabled = false,
            SetupInstructions = new[]
            {
                "Open Excel",
                "Go to File → Options → Trust Center",
                "Click 'Trust Center Settings'",
                "Select 'Macro Settings'",
                "Check '✓ Trust access to the VBA project object model'",
                "Click OK twice to save settings"
            },
            DocumentationUrl = "https://support.microsoft.com/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6",
            Explanation = "VBA operations require 'Trust access to the VBA project object model' to be enabled in Excel settings."
        };

        // Assert - Verify all properties are accessible and have expected values
        Assert.False(trustResult.Success);
        Assert.Equal("VBA trust access is not enabled", trustResult.ErrorMessage);
        Assert.False(trustResult.IsTrustEnabled);
        Assert.NotNull(trustResult.SetupInstructions);
        Assert.Equal(6, trustResult.SetupInstructions.Length);
        Assert.Contains("Open Excel", trustResult.SetupInstructions);
        Assert.False(string.IsNullOrEmpty(trustResult.DocumentationUrl));
        Assert.False(string.IsNullOrEmpty(trustResult.Explanation));
    }

    [Fact]
    public void TestVbaTrustScope_EnablesAndDisablesTrust()
    {
        // Arrange - Check initial trust state
        bool initialTrustState = IsVbaTrustEnabled();

        // Act - Use TestVbaTrustScope
        using (new TestVbaTrustScope())
        {
            // Inside the scope, VBA trust should be enabled
            Assert.True(IsVbaTrustEnabled(), "VBA trust should be enabled inside TestVbaTrustScope");
        }

        // Assert - After scope disposal, trust should be restored to initial state
        bool finalTrustState = IsVbaTrustEnabled();
        Assert.Equal(initialTrustState, finalTrustState);
    }
}
