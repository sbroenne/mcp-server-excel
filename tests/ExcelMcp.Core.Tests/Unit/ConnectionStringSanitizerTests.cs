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
    [Fact]
    public void Sanitize_RemovesCredentialValues()
    {
        var cases = new[]
        {
            ("Pass" + "word=sensitive-password;Server=localhost", "sensitive-password", "Server=localhost"),
            ("p" + "wd=\"sensitive;value\";Server=localhost", "sensitive;value", "Server=localhost"),
            ("User ID='sensitive-user';Server=localhost", "sensitive-user", "Server=localhost"),
            ("uid={sensitive-user};Server=localhost", "sensitive-user", "Server=localhost"),
            ("Username=sensitive-user;Server=localhost", "sensitive-user", "Server=localhost"),
            ("User=sensitive-user;Server=localhost", "sensitive-user", "Server=localhost"),
            ("Client Secret=sensitive-client;Server=localhost", "sensitive-client", "Server=localhost"),
            ("Access_Token=sensitive-token;Server=localhost", "sensitive-token", "Server=localhost"),
            ("ApiKey=sensitive-key;Server=localhost", "sensitive-key", "Server=localhost"),
            ("Account Key=sensitive-key;Server=localhost", "sensitive-key", "Server=localhost"),
            ("Data Source=https://" + "sensitive-user:sensitive-password@example.test/data", "sensitive-password", "example.test/data"),
            ("Data Source=https://" + "sensitive-user:sensitive-password@example.test/data", "sensitive-user", "example.test/data"),
            ("Data Source=https://example.test/data?access_" + "token=sensitive-token&mode=read", "sensitive-token", "mode=read"),
            ("Data Source=https://example.test/data?client-secret=sensitive-client&mode=read", "sensitive-client", "mode=read"),
            ("Data Source=https://example.test/data?sig=sensitive-signature&mode=read", "sensitive-signature", "mode=read")
        };

        foreach (var (connectionString, sensitiveValue, expectedProperty) in cases)
        {
            var sanitized = ConnectionStringSanitizer.Sanitize(connectionString);

            Assert.NotNull(sanitized);
            Assert.Contains("(redacted)", sanitized, StringComparison.Ordinal);
            Assert.DoesNotContain(sensitiveValue, sanitized, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(expectedProperty, sanitized, StringComparison.Ordinal);
        }
    }
}
