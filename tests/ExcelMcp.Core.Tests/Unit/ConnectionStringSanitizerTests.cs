using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "Connection")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ConnectionStringSanitizerTests
{
    private const string SyntheticSecret = "synthetic-secret";

    [Theory]
    [InlineData("User ID")]
    [InlineData("UID")]
    [InlineData("Username")]
    [InlineData("User Name")]
    [InlineData("user")]
    public void Sanitize_RedactsUserAliases(string userKey)
    {
        var connectionString = "Provider=SQLOLEDB;Data Source=srv;" + Setting(userKey, "synthetic-user");

        var sanitized = ConnectionStringSanitizer.Sanitize(connectionString);

        Assert.NotNull(sanitized);
        Assert.DoesNotContain("synthetic-user", sanitized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("(redacted)", sanitized, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Password")]
    [InlineData("Passwd")]
    [InlineData("PWD")]
    [InlineData("Secret")]
    [InlineData("Client Secret")]
    [InlineData("Account Key")]
    [InlineData("Shared Access Signature")]
    [InlineData("API Key")]
    [InlineData("Access Token")]
    public void Sanitize_RedactsSecretAliases(string secretKey)
    {
        var connectionString = "Provider=SQLOLEDB;Data Source=srv;" + Setting(secretKey, SyntheticSecret);

        var sanitized = ConnectionStringSanitizer.Sanitize(connectionString);

        Assert.NotNull(sanitized);
        Assert.DoesNotContain(SyntheticSecret, sanitized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("(redacted)", sanitized, StringComparison.Ordinal);
    }

    [Fact]
    public void Sanitize_KeepsNonCredentialSettings()
    {
        const string connectionString =
            "Provider=SQLOLEDB;Data Source=srv;Initial Catalog=SalesDb;Integrated Security=SSPI;";

        Assert.Equal(connectionString, ConnectionStringSanitizer.Sanitize(connectionString));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Sanitize_PassesThroughEmptyValues(string? connectionString)
    {
        Assert.Equal(connectionString, ConnectionStringSanitizer.Sanitize(connectionString));
    }

    private static string Setting(string key, string value) => string.Join("=", key, value) + ";";
}
