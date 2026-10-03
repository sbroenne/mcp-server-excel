using Sbroenne.ExcelMcp.Core.DataModel;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DataModelRefreshErrorMessageTests
{
    [Theory]
    [InlineData("DataSource.Error: synthetic source failure")]
    [InlineData("Synthetic Excel engine error")]
    public void RefreshFailed_PreservesDetailsWithoutInventingCapabilityFailure(string details)
    {
        var message = DataModelErrorMessages.RefreshFailed(details);
        Assert.Contains(details, message);
        Assert.Contains("refresh failed", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("not supported", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tables individually", message, StringComparison.OrdinalIgnoreCase);
    }
}
