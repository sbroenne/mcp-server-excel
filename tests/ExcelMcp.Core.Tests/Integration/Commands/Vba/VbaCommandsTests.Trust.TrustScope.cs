using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.VbaTrust;

/// <summary>
/// Tests for VbaTrustRequiredResult model
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
            Explanation = "VBA operations require 'Trust access to the VBA project object model' to be enabled in Excel settings."
        };

        // Assert - Verify all properties are accessible and have expected values
        Assert.False(trustResult.Success);
        Assert.Equal("VBA trust access is not enabled", trustResult.ErrorMessage);
        Assert.False(trustResult.IsTrustEnabled);
        Assert.False(string.IsNullOrEmpty(trustResult.Explanation));
    }
}
