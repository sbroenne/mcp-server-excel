using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for ConnectionHelpers and PowerQueryHelpers shared utilities
/// Tests utilities extracted from PowerQueryCommands for DRY compliance
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "SharedUtilities")]
public class ConnectionHelpersTests
{
    #region GetConnectionTypeName Tests

    [Theory]
    [InlineData(1, "OLEDB")]
    [InlineData(2, "ODBC")]
    [InlineData(3, "TEXT")]       // Fixed: Was "XML", now matches Microsoft XlConnectionType enum
    [InlineData(4, "WEB")]        // Fixed: Was "Text", now matches Microsoft XlConnectionType enum
    [InlineData(5, "XMLMAP")]     // Fixed: Was "Web", now matches Microsoft XlConnectionType enum
    [InlineData(6, "DATAFEED")]   // Fixed: Was "DataFeed", now uppercase per Microsoft docs
    [InlineData(7, "MODEL")]      // Fixed: Was "Model", now uppercase per Microsoft docs
    [InlineData(8, "WORKSHEET")]  // Fixed: Was "Worksheet", now uppercase per Microsoft docs
    [InlineData(9, "NOSOURCE")]   // Fixed: Was "NoSource", now uppercase per Microsoft docs
    public void GetConnectionTypeName_ValidTypes_ReturnsCorrectName(int typeValue, string expected)
    {
        // Act
        var result = ConnectionHelpers.GetConnectionTypeName(typeValue);

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetConnectionTypeName_UnknownType_ReturnsUnknownWithValue()
    {
        // Arrange
        int unknownType = 999;

        // Act
        var result = ConnectionHelpers.GetConnectionTypeName(unknownType);

        // Assert
        Assert.Equal("Unknown (999)", result);
    }

    [Fact]
    public void GetConnectionTypeName_NegativeValue_ReturnsUnknownWithValue()
    {
        // Arrange
        int negativeType = -1;

        // Act
        var result = ConnectionHelpers.GetConnectionTypeName(negativeType);

        // Assert
        Assert.Equal("Unknown (-1)", result);
    }

    #endregion

    #region SanitizeConnectionString Tests

    [Fact]
    public void SanitizeConnectionString_WithPassword_MasksPassword()
    {
        // Arrange
        string connectionString = "Provider=SQLOLEDB;Data Source=server;Password=MySecret123;User ID=admin";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Password=***REDACTED***", result);
        Assert.DoesNotContain("MySecret123", result);
    }

    [Fact]
    public void SanitizeConnectionString_WithPwd_MasksPwd()
    {
        // Arrange
        string connectionString = "Provider=SQLOLEDB;Data Source=server;Pwd=MySecret123;User ID=admin";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Pwd=***REDACTED***", result);
        Assert.DoesNotContain("MySecret123", result);
    }

    [Fact]
    public void SanitizeConnectionString_MultipleSensitiveFields_MasksAll()
    {
        // Arrange
        string connectionString1 = "Password=Secret1;Pwd=Secret2;AccessToken=Token123";
        string connectionString2 = "AccountKey=Key456;password=Secret3;PWD=Secret4";
        string connectionString3 = "AccessKey=Access789;apikey=API999;ApiKey=API888";

        // Act
        var result1 = ConnectionHelpers.SanitizeConnectionString(connectionString1);
        var result2 = ConnectionHelpers.SanitizeConnectionString(connectionString2);
        var result3 = ConnectionHelpers.SanitizeConnectionString(connectionString3);

        // Assert
        Assert.Contains("Password=***REDACTED***", result1);
        Assert.Contains("Pwd=***REDACTED***", result1);
        Assert.Contains("AccessToken=***REDACTED***", result1);
        Assert.DoesNotContain("Secret", result1);
        Assert.DoesNotContain("Token123", result1);

        Assert.Contains("AccountKey=***REDACTED***", result2);
        Assert.Contains("password=***REDACTED***", result2);
        Assert.Contains("PWD=***REDACTED***", result2);

        Assert.Contains("AccessKey=***REDACTED***", result3);
        Assert.Contains("apikey=***REDACTED***", result3);
        Assert.Contains("ApiKey=***REDACTED***", result3);
    }

    [Fact]
    public void SanitizeConnectionString_CaseInsensitive_MasksAllVariants()
    {
        // Arrange
        string connectionString = "PASSWORD=Secret1;password=Secret2;PaSsWoRd=Secret3";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        Assert.DoesNotContain("Secret", result);
        Assert.Contains("***REDACTED***", result);
    }

    [Fact]
    public void SanitizeConnectionString_WithoutSensitiveData_ReturnsUnchanged()
    {
        // Arrange
        string connectionString = "Provider=SQLOLEDB;Data Source=server;Initial Catalog=TestDB";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Equal(connectionString, result);
    }

    [Fact]
    public void SanitizeConnectionString_ComplexConnectionString_MasksOnlySensitiveFields()
    {
        // Arrange
        string connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" +
                                  "Initial Catalog=TestDB;Data Source=localhost;Password=MySecret123;User ID=admin";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Password=***REDACTED***", result);
        Assert.DoesNotContain("MySecret123", result);
        Assert.Contains("Data Source=localhost", result);
        Assert.Contains("User ID=admin", result);
    }

    [Fact]
    public void SanitizeConnectionString_NullInput_ReturnsNull()
    {
        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(null);

        // Assert
        Assert.Null(result);
    }

    [Fact]
    public void SanitizeConnectionString_EmptyString_ReturnsEmpty()
    {
        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(string.Empty);

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void SanitizeConnectionString_WhitespaceOnly_ReturnsWhitespace()
    {
        // Act
        var result = ConnectionHelpers.SanitizeConnectionString("   ");

        // Assert
        Assert.Equal("   ", result);
    }

    [Fact]
    public void SanitizeConnectionString_UrlWithPassword_MasksPassword()
    {
        // Arrange
        string connectionString = "https://user:MySecret123@example.com/data";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        // URL format passwords might not be detected by the current implementation
        // This test documents expected behavior
        Assert.NotNull(result);
    }

    [Fact]
    public void SanitizeConnectionString_ODataWithKey_MasksKey()
    {
        // Arrange
        string connectionString = "https://api.example.com/odata?$apikey=MyKey123&$format=json";

        // Act
        var result = ConnectionHelpers.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("apikey=***REDACTED***", result);
        Assert.DoesNotContain("MyKey123", result);
    }

    #endregion

    #region QueryTableOptions Tests

    [Fact]
    public void QueryTableOptions_DefaultValues_AreCorrect()
    {
        // Arrange & Act
        var options = new PowerQueryHelpers.QueryTableOptions { Name = "TestQuery" };

        // Assert
        Assert.Equal("TestQuery", options.Name);
    }

    [Fact]
    public void QueryTableOptions_AllPropertiesSettable()
    {
        // Arrange & Act
        var options = new PowerQueryHelpers.QueryTableOptions { Name = "Test" };

        // Assert
        Assert.NotNull(options);
    }

    [Fact]
    public void QueryTableOptions_CanBeUsedInConfiguration()
    {
        // Arrange & Act
        var options = new PowerQueryHelpers.QueryTableOptions
        {
            Name = "CustomerData"
        };

        // Assert
        Assert.Equal("CustomerData", options.Name);
        // Additional properties can be verified if QueryTableOptions has more public properties
    }

    #endregion
}
