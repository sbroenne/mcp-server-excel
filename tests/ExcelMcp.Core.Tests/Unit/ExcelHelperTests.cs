using Sbroenne.ExcelMcp.Core;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for ExcelHelper shared utilities (Phase 0 - Connection Management)
/// Tests utilities extracted from PowerQueryCommands for DRY compliance
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "SharedUtilities")]
public class ExcelHelperTests
{
    #region GetConnectionTypeName Tests

    [Theory]
    [InlineData(1, "OLEDB")]
    [InlineData(2, "ODBC")]
    [InlineData(3, "XML")]
    [InlineData(4, "Text")]
    [InlineData(5, "Web")]
    [InlineData(6, "DataFeed")]
    [InlineData(7, "Model")]
    [InlineData(8, "Worksheet")]
    [InlineData(9, "NoSource")]
    public void GetConnectionTypeName_ValidTypes_ReturnsCorrectName(int typeValue, string expected)
    {
        // Act
        var result = ExcelHelper.GetConnectionTypeName(typeValue);

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetConnectionTypeName_UnknownType_ReturnsUnknownWithValue()
    {
        // Arrange
        int unknownType = 999;

        // Act
        var result = ExcelHelper.GetConnectionTypeName(unknownType);

        // Assert
        Assert.Equal("Unknown (999)", result);
    }

    [Fact]
    public void GetConnectionTypeName_NegativeValue_ReturnsUnknownWithValue()
    {
        // Arrange
        int negativeType = -1;

        // Act
        var result = ExcelHelper.GetConnectionTypeName(negativeType);

        // Assert
        Assert.Equal("Unknown (-1)", result);
    }

    #endregion

    #region SanitizeConnectionString Tests

    [Fact]
    public void SanitizeConnectionString_WithPassword_MasksPassword()
    {
        // Arrange
        string connectionString = "Server=myserver;Database=mydb;Password=SecretP@ssw0rd;Trusted_Connection=False";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Password=***", result);
        Assert.DoesNotContain("SecretP@ssw0rd", result);
        Assert.Contains("Server=myserver", result);
        Assert.Contains("Database=mydb", result);
    }

    [Fact]
    public void SanitizeConnectionString_WithPwd_MasksPwd()
    {
        // Arrange
        string connectionString = "Server=myserver;Database=mydb;Pwd=MySecret123;Trusted_Connection=False";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Pwd=***", result);
        Assert.DoesNotContain("MySecret123", result);
        Assert.Contains("Server=myserver", result);
    }

    [Fact]
    public void SanitizeConnectionString_CaseInsensitive_MasksPasswordRegardlessOfCase()
    {
        // Arrange
        string connectionString1 = "Server=myserver;PASSWORD=Secret123;Database=mydb";
        string connectionString2 = "Server=myserver;password=Secret456;Database=mydb";
        string connectionString3 = "Server=myserver;PaSsWoRd=Secret789;Database=mydb";

        // Act
        var result1 = ExcelHelper.SanitizeConnectionString(connectionString1);
        var result2 = ExcelHelper.SanitizeConnectionString(connectionString2);
        var result3 = ExcelHelper.SanitizeConnectionString(connectionString3);

        // Assert
        Assert.All(new[] { result1, result2, result3 }, r =>
        {
            Assert.Contains("***", r);
            Assert.DoesNotContain("Secret", r);
        });
    }

    [Fact]
    public void SanitizeConnectionString_WithSpaces_HandlesSpacesAroundEquals()
    {
        // Arrange
        string connectionString = "Server=myserver;Password = SecretWithSpaces;Database=mydb";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert - Regex normalizes spaces, so "Password = " becomes "Password="
        Assert.Contains("Password=***", result);
        Assert.DoesNotContain("SecretWithSpaces", result);
    }

    [Fact]
    public void SanitizeConnectionString_MultiplePasswords_MasksAllPasswords()
    {
        // Arrange
        string connectionString = "Server=myserver;Password=Secret1;Database=mydb;Pwd=Secret2";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Password=***", result);
        Assert.Contains("Pwd=***", result);
        Assert.DoesNotContain("Secret1", result);
        Assert.DoesNotContain("Secret2", result);
    }

    [Fact]
    public void SanitizeConnectionString_NoPassword_ReturnsUnchanged()
    {
        // Arrange
        string connectionString = "Server=myserver;Database=mydb;Trusted_Connection=True";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Equal(connectionString, result);
    }

    [Fact]
    public void SanitizeConnectionString_Null_ReturnsEmptyString()
    {
        // Act
        var result = ExcelHelper.SanitizeConnectionString(null);

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void SanitizeConnectionString_EmptyString_ReturnsEmptyString()
    {
        // Act
        var result = ExcelHelper.SanitizeConnectionString(string.Empty);

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void SanitizeConnectionString_WhitespaceOnly_ReturnsEmptyString()
    {
        // Act
        var result = ExcelHelper.SanitizeConnectionString("   ");

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void SanitizeConnectionString_PasswordAtEnd_MasksPassword()
    {
        // Arrange
        string connectionString = "Server=myserver;Database=mydb;Password=EndSecret";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Password=***", result);
        Assert.DoesNotContain("EndSecret", result);
    }

    [Fact]
    public void SanitizeConnectionString_ComplexOLEDBConnection_MasksPasswordProperly()
    {
        // Arrange
        string connectionString = "Provider=SQLOLEDB;Data Source=myserver;Initial Catalog=mydb;User ID=myuser;Password=MyComplexP@ssw0rd!;Persist Security Info=False";

        // Act
        var result = ExcelHelper.SanitizeConnectionString(connectionString);

        // Assert
        Assert.Contains("Password=***", result);
        Assert.DoesNotContain("MyComplexP@ssw0rd!", result);
        Assert.Contains("Provider=SQLOLEDB", result);
        Assert.Contains("User ID=myuser", result);
    }

    #endregion

    #region QueryTableOptions Tests

    [Fact]
    public void QueryTableOptions_RequiredName_CanBeSet()
    {
        // Act
        var options = new ExcelHelper.QueryTableOptions { Name = "TestQuery" };

        // Assert
        Assert.Equal("TestQuery", options.Name);
    }

    [Fact]
    public void QueryTableOptions_DefaultValues_AreCorrect()
    {
        // Act
        var options = new ExcelHelper.QueryTableOptions { Name = "Test" };

        // Assert
        Assert.False(options.BackgroundQuery);
        Assert.False(options.RefreshOnFileOpen);
        Assert.False(options.SavePassword);
        Assert.True(options.PreserveColumnInfo);
        Assert.True(options.PreserveFormatting);
        Assert.True(options.AdjustColumnWidth);
        Assert.False(options.RefreshImmediately);
    }

    [Fact]
    public void QueryTableOptions_AllProperties_CanBeSet()
    {
        // Act
        var options = new ExcelHelper.QueryTableOptions
        {
            Name = "MyQuery",
            BackgroundQuery = true,
            RefreshOnFileOpen = true,
            SavePassword = true,
            PreserveColumnInfo = false,
            PreserveFormatting = false,
            AdjustColumnWidth = false,
            RefreshImmediately = true
        };

        // Assert
        Assert.Equal("MyQuery", options.Name);
        Assert.True(options.BackgroundQuery);
        Assert.True(options.RefreshOnFileOpen);
        Assert.True(options.SavePassword);
        Assert.False(options.PreserveColumnInfo);
        Assert.False(options.PreserveFormatting);
        Assert.False(options.AdjustColumnWidth);
        Assert.True(options.RefreshImmediately);
    }

    #endregion
}
