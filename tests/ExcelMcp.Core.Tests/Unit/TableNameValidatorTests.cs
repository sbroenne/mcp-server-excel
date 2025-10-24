using Sbroenne.ExcelMcp.Core.Security;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class TableNameValidatorTests
{
    #region Valid Names

    [Fact]
    public void ValidateTableName_ValidName_ReturnsName()
    {
        // Valid simple name
        string result = TableNameValidator.ValidateTableName("MyTable");
        Assert.Equal("MyTable", result);
    }

    [Fact]
    public void ValidateTableName_WithUnderscore_ReturnsName()
    {
        string result = TableNameValidator.ValidateTableName("My_Table");
        Assert.Equal("My_Table", result);
    }

    [Fact]
    public void ValidateTableName_WithPeriod_ReturnsName()
    {
        string result = TableNameValidator.ValidateTableName("My.Table");
        Assert.Equal("My.Table", result);
    }

    [Fact]
    public void ValidateTableName_WithNumbers_ReturnsName()
    {
        string result = TableNameValidator.ValidateTableName("Table123");
        Assert.Equal("Table123", result);
    }

    [Fact]
    public void ValidateTableName_StartsWithUnderscore_ReturnsName()
    {
        string result = TableNameValidator.ValidateTableName("_MyTable");
        Assert.Equal("_MyTable", result);
    }

    [Fact]
    public void ValidateTableName_TrimsWhitespace_ReturnsName()
    {
        string result = TableNameValidator.ValidateTableName("  MyTable  ");
        Assert.Equal("MyTable", result);
    }

    #endregion

    #region Null/Empty Validation

    [Fact]
    public void ValidateTableName_Null_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName(null!));
        Assert.Contains("cannot be null", ex.Message);
    }

    [Fact]
    public void ValidateTableName_Empty_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName(""));
        Assert.Contains("cannot be null", ex.Message);
    }

    [Fact]
    public void ValidateTableName_WhitespaceOnly_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("   "));
        Assert.Contains("cannot be null", ex.Message);
    }

    #endregion

    #region Length Validation

    [Fact]
    public void ValidateTableName_ExactlyMaxLength_ReturnsName()
    {
        string longName = new string('A', 255);
        string result = TableNameValidator.ValidateTableName(longName);
        Assert.Equal(longName, result);
    }

    [Fact]
    public void ValidateTableName_TooLong_ThrowsArgumentException()
    {
        string tooLong = new string('A', 256);
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName(tooLong));
        Assert.Contains("too long", ex.Message);
        Assert.Contains("256 characters", ex.Message);
    }

    #endregion

    #region Space Validation

    [Fact]
    public void ValidateTableName_WithSpace_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("Invalid Name"));
        Assert.Contains("cannot contain spaces", ex.Message);
    }

    [Fact]
    public void ValidateTableName_WithMultipleSpaces_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("My Table Name"));
        Assert.Contains("cannot contain spaces", ex.Message);
    }

    #endregion

    #region First Character Validation

    [Fact]
    public void ValidateTableName_StartsWithNumber_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("123Table"));
        Assert.Contains("must start with a letter or underscore", ex.Message);
        Assert.Contains("'1'", ex.Message);
    }

    [Fact]
    public void ValidateTableName_StartsWithPeriod_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName(".Table"));
        Assert.Contains("must start with a letter or underscore", ex.Message);
    }

    [Fact]
    public void ValidateTableName_StartsWithSpecialChar_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("@Table"));
        Assert.Contains("must start with a letter or underscore", ex.Message);
    }

    #endregion

    #region Invalid Character Validation

    [Fact]
    public void ValidateTableName_WithAtSign_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("My@Table"));
        Assert.Contains("invalid character: '@'", ex.Message);
    }

    [Fact]
    public void ValidateTableName_WithHyphen_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("My-Table"));
        Assert.Contains("invalid character: '-'", ex.Message);
    }

    [Fact]
    public void ValidateTableName_WithDollar_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("My$Table"));
        Assert.Contains("invalid character: '$'", ex.Message);
    }

    [Fact]
    public void ValidateTableName_WithParenthesis_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("My(Table)"));
        Assert.Contains("invalid character", ex.Message);
    }

    #endregion

    #region Reserved Names Validation

    [Fact]
    public void ValidateTableName_PrintArea_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("Print_Area"));
        Assert.Contains("reserved name", ex.Message);
    }

    [Fact]
    public void ValidateTableName_PrintTitles_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("Print_Titles"));
        Assert.Contains("reserved name", ex.Message);
    }

    [Fact]
    public void ValidateTableName_FilterDatabase_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("_FilterDatabase"));
        Assert.Contains("reserved name", ex.Message);
    }

    [Fact]
    public void ValidateTableName_ReservedName_CaseInsensitive_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("PRINT_AREA"));
        Assert.Contains("reserved name", ex.Message);
    }

    #endregion

    #region Cell Reference Validation

    [Fact]
    public void ValidateTableName_A1Reference_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("A1"));
        Assert.Contains("looks like a cell reference", ex.Message);
    }

    [Fact]
    public void ValidateTableName_Z100Reference_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("Z100"));
        Assert.Contains("looks like a cell reference", ex.Message);
    }

    [Fact]
    public void ValidateTableName_R1C1Reference_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("R1C1"));
        Assert.Contains("looks like a cell reference", ex.Message);
    }

    [Fact]
    public void ValidateTableName_R10C5Reference_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("R10C5"));
        Assert.Contains("looks like a cell reference", ex.Message);
    }

    [Fact]
    public void ValidateTableName_AB123_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName("AB123"));
        Assert.Contains("looks like a cell reference", ex.Message);
    }

    [Fact]
    public void ValidateTableName_NotCellReference_DoesNotThrow()
    {
        // These should NOT be detected as cell references
        string result1 = TableNameValidator.ValidateTableName("Table_A1");
        string result2 = TableNameValidator.ValidateTableName("A1_Table");
        string result3 = TableNameValidator.ValidateTableName("MyR1C1Data");
        
        Assert.Equal("Table_A1", result1);
        Assert.Equal("A1_Table", result2);
        Assert.Equal("MyR1C1Data", result3);
    }

    #endregion

    #region TryValidateTableName

    [Fact]
    public void TryValidateTableName_ValidName_ReturnsTrue()
    {
        var (isValid, errorMessage) = TableNameValidator.TryValidateTableName("MyTable");
        
        Assert.True(isValid);
        Assert.Null(errorMessage);
    }

    [Fact]
    public void TryValidateTableName_InvalidName_ReturnsFalseWithMessage()
    {
        var (isValid, errorMessage) = TableNameValidator.TryValidateTableName("Invalid Name");
        
        Assert.False(isValid);
        Assert.NotNull(errorMessage);
        Assert.Contains("cannot contain spaces", errorMessage);
    }

    [Fact]
    public void TryValidateTableName_NullName_ReturnsFalseWithMessage()
    {
        var (isValid, errorMessage) = TableNameValidator.TryValidateTableName(null!);
        
        Assert.False(isValid);
        Assert.NotNull(errorMessage);
        Assert.Contains("cannot be null", errorMessage);
    }

    [Fact]
    public void TryValidateTableName_ReservedName_ReturnsFalseWithMessage()
    {
        var (isValid, errorMessage) = TableNameValidator.TryValidateTableName("Print_Area");
        
        Assert.False(isValid);
        Assert.NotNull(errorMessage);
        Assert.Contains("reserved name", errorMessage);
    }

    #endregion

    #region Security - Injection Prevention

    [Theory]
    [InlineData("=SUM(A1:A10)")] // Formula injection attempt
    [InlineData("+1-1")] // Formula injection
    [InlineData("@Table")] // Special character
    [InlineData("Table;DROP")] // SQL-style injection
    [InlineData("Table|cmd")] // Command injection
    [InlineData("Table&calc")] // Command injection
    public void ValidateTableName_InjectionAttempts_ThrowsArgumentException(string maliciousName)
    {
        Assert.Throws<ArgumentException>(() => 
            TableNameValidator.ValidateTableName(maliciousName));
    }

    #endregion
}
