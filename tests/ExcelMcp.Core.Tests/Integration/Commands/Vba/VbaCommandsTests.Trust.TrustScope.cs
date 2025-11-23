using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

/// <summary>
/// Placeholder for trust guidance tests. Trust checks now throw InvalidOperationException directly.
/// </summary>
public partial class VbaCommandsTests
{
    [Fact(Skip = "Trust guidance now enforced via InvalidOperationException; no Result object to verify.")]
    public void TrustGuidance_ReplacedByExceptions()
    {
    }
}
