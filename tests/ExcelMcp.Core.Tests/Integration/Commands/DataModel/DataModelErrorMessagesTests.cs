using Sbroenne.ExcelMcp.Core.DataModel;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Fast")]
public sealed class DataModelErrorMessagesTests
{
    [Fact]
    public void MsolapClassNotRegistered_IncludesProviderDiagnosticsWithoutClaimingProviderIsAbsent()
    {
        var diagnostics = new DataModelAdoDiagnostics
        {
            ProviderName = "MSOLAP.5",
            ConnectionString = "Provider=MSOLAP.5;Data Source=$Embedded$;Persist Security Info=True;User ID=admin;Password=secret"
        };

        var message = DataModelErrorMessages.MsolapClassNotRegistered(diagnostics);

        Assert.Contains("MSOLAP.5", message, StringComparison.Ordinal);
        Assert.Contains("class is not registered", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Excel Data Model ADO connection", message, StringComparison.Ordinal);
        Assert.DoesNotContain("which is not installed", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("admin", message, StringComparison.OrdinalIgnoreCase);
    }
}
