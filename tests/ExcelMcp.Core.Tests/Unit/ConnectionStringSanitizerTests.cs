using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ConnectionStringSanitizerTests
{
    [Theory]
    [InlineData("Password=secret-value;Server=localhost")]
    [InlineData("PWD=\"secret;value\";Server=localhost")]
    [InlineData("User ID='secret-user';Server=localhost")]
    [InlineData("uid={secret-user};Server=localhost")]
    public void Sanitize_RemovesCredentialValues(string connectionString)
    {
        var sanitized = ConnectionStringSanitizer.Sanitize(connectionString);

        Assert.NotNull(sanitized);
        Assert.Contains("(redacted)", sanitized, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", sanitized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Server=localhost", sanitized, StringComparison.Ordinal);
    }
}
