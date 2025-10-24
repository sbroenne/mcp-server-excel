using Sbroenne.ExcelMcp.Core.Security;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class RangeValidatorTests
{
    #region ValidateRangeAddress Tests

    [Fact]
    public void ValidateRangeAddress_ValidSingleCell_ReturnsAddress()
    {
        string result = RangeValidator.ValidateRangeAddress("A1");
        Assert.Equal("A1", result);
    }

    [Fact]
    public void ValidateRangeAddress_ValidRange_ReturnsAddress()
    {
        string result = RangeValidator.ValidateRangeAddress("A1:B10");
        Assert.Equal("A1:B10", result);
    }

    [Fact]
    public void ValidateRangeAddress_WithAbsoluteReference_ReturnsAddress()
    {
        string result = RangeValidator.ValidateRangeAddress("$A$1:$B$10");
        Assert.Equal("$A$1:$B$10", result);
    }

    [Fact]
    public void ValidateRangeAddress_WithSheetName_ReturnsAddress()
    {
        string result = RangeValidator.ValidateRangeAddress("Sheet1!A1:B10");
        Assert.Equal("Sheet1!A1:B10", result);
    }

    [Fact]
    public void ValidateRangeAddress_TrimsWhitespace_ReturnsAddress()
    {
        string result = RangeValidator.ValidateRangeAddress("  A1:B10  ");
        Assert.Equal("A1:B10", result);
    }

    [Fact]
    public void ValidateRangeAddress_Null_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            RangeValidator.ValidateRangeAddress(null!));
        Assert.Contains("cannot be null", ex.Message);
    }

    [Fact]
    public void ValidateRangeAddress_Empty_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            RangeValidator.ValidateRangeAddress(""));
        Assert.Contains("cannot be null", ex.Message);
    }

    [Fact]
    public void ValidateRangeAddress_Whitespace_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            RangeValidator.ValidateRangeAddress("   "));
        Assert.Contains("cannot be null", ex.Message);
    }

    [Fact]
    public void ValidateRangeAddress_TooLong_ThrowsArgumentException()
    {
        string tooLong = new string('A', 256);
        var ex = Assert.Throws<ArgumentException>(() =>
            RangeValidator.ValidateRangeAddress(tooLong));
        Assert.Contains("too long", ex.Message);
    }

    [Theory]
    [InlineData("A1;B2")] // Semicolon
    [InlineData("A1,B2")] // Comma (though valid in some contexts)
    [InlineData("A1@B2")] // At sign
    [InlineData("A1#B2")] // Hash
    [InlineData("A1%B2")] // Percent
    [InlineData("A1&B2")] // Ampersand
    [InlineData("A1*B2")] // Asterisk
    [InlineData("A1(B2)")] // Parentheses
    [InlineData("A1[B2]")] // Brackets
    [InlineData("A1{B2}")] // Braces
    public void ValidateRangeAddress_InvalidCharacters_ThrowsArgumentException(string invalidAddress)
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            RangeValidator.ValidateRangeAddress(invalidAddress));
        Assert.Contains("invalid character", ex.Message);
    }

    #endregion

    #region Mock Range Object Helper

    /// <summary>
    /// Mock range object for testing purposes
    /// </summary>
    public class MockRange
    {
        public MockRows Rows { get; set; } = new MockRows();
        public MockColumns Columns { get; set; } = new MockColumns();

        public static MockRange Create(int rows, int cols)
        {
            return new MockRange
            {
                Rows = new MockRows { Count = rows },
                Columns = new MockColumns { Count = cols }
            };
        }
    }

    public class MockRows
    {
        public int Count { get; set; }
    }

    public class MockColumns
    {
        public int Count { get; set; }
    }

    #endregion

    #region ValidateRange Tests - Basic Validation

    [Fact]
    public void ValidateRange_ValidSmallRange_DoesNotThrow()
    {
        dynamic range = MockRange.Create(10, 10);
        
        // Should not throw
        RangeValidator.ValidateRange(range);
    }

    [Fact]
    public void ValidateRange_ValidLargeRange_DoesNotThrow()
    {
        // 1000 x 1000 = 1,000,000 cells (exactly at limit)
        dynamic range = MockRange.Create(1000, 1000);
        
        // Should not throw
        RangeValidator.ValidateRange(range);
    }

    [Fact]
    public void ValidateRange_NullRange_ThrowsArgumentNullException()
    {
        var ex = Assert.Throws<ArgumentNullException>(() =>
            RangeValidator.ValidateRange(null!));
        Assert.Contains("cannot be null", ex.Message);
    }

    [Fact]
    public void ValidateRange_ZeroRows_ThrowsArgumentException()
    {
        dynamic range = MockRange.Create(0, 10);
        
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range);
        });
        Assert.Contains("at least one cell", ex.Message);
    }

    [Fact]
    public void ValidateRange_ZeroColumns_ThrowsArgumentException()
    {
        dynamic range = MockRange.Create(10, 0);
        
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range);
        });
        Assert.Contains("at least one cell", ex.Message);
    }

    [Fact]
    public void ValidateRange_NegativeRows_ThrowsArgumentException()
    {
        dynamic range = MockRange.Create(-1, 10);
        
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range);
        });
        Assert.Contains("at least one cell", ex.Message);
    }

    #endregion

    #region ValidateRange Tests - DoS Prevention

    [Fact]
    public void ValidateRange_ExceedsDefaultLimit_ThrowsArgumentException()
    {
        // 1001 x 1000 = 1,001,000 cells (just over limit)
        dynamic range = MockRange.Create(1001, 1000);
        
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range);
        });
        Assert.Contains("too large", ex.Message);
        Assert.Contains("1,001,000", ex.Message);
        Assert.Contains("denial-of-service", ex.Message);
    }

    [Fact]
    public void ValidateRange_ExtremelyLarge_ThrowsArgumentException()
    {
        // 100,000 x 100 = 10,000,000 cells (way over limit)
        dynamic range = MockRange.Create(100000, 100);
        
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range);
        });
        Assert.Contains("too large", ex.Message);
        Assert.Contains("10,000,000", ex.Message);
    }

    [Fact]
    public void ValidateRange_CustomMaxCells_RespectedLimit()
    {
        dynamic range = MockRange.Create(100, 100); // 10,000 cells
        
        // Custom limit of 5,000 cells
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range, maxCells: 5000);
        });
        Assert.Contains("too large", ex.Message);
        Assert.Contains("10,000", ex.Message);
        Assert.Contains("5,000", ex.Message);
    }

    [Fact]
    public void ValidateRange_CustomMaxCells_AllowsLargerRange()
    {
        dynamic range = MockRange.Create(2000, 1000); // 2,000,000 cells
        
        // Custom limit of 3,000,000 cells - should not throw
        RangeValidator.ValidateRange(range, maxCells: 3_000_000);
    }

    #endregion

    #region TryValidateRange Tests

    [Fact]
    public void TryValidateRange_ValidRange_ReturnsTrue()
    {
        dynamic range = MockRange.Create(10, 20);
        
        var result = RangeValidator.TryValidateRange(range);
        
        Assert.True(result.Item1); // isValid
        Assert.Null(result.Item2); // errorMessage
        Assert.Equal(10, result.Item3); // rowCount
        Assert.Equal(20, result.Item4); // colCount
        Assert.Equal(200, result.Item5); // cellCount
    }

    [Fact]
    public void TryValidateRange_NullRange_ReturnsFalse()
    {
        var result = RangeValidator.TryValidateRange(null!);
        
        Assert.False(result.Item1); // isValid
        Assert.NotNull(result.Item2); // errorMessage
        Assert.Contains("null", result.Item2);
        Assert.Equal(0, result.Item3); // rowCount
        Assert.Equal(0, result.Item4); // colCount
        Assert.Equal(0, result.Item5); // cellCount
    }

    [Fact]
    public void TryValidateRange_TooLarge_ReturnsFalse()
    {
        dynamic range = MockRange.Create(2000, 1000); // 2,000,000 cells
        
        var result = RangeValidator.TryValidateRange(range);
        
        Assert.False(result.Item1);
        Assert.NotNull(result.Item2);
        Assert.Contains("too large", result.Item2);
        Assert.Equal(2000, result.Item3);
        Assert.Equal(1000, result.Item4);
        Assert.Equal(2_000_000, result.Item5);
    }

    [Fact]
    public void TryValidateRange_ZeroDimensions_ReturnsFalse()
    {
        dynamic range = MockRange.Create(0, 10);
        
        var result = RangeValidator.TryValidateRange(range);
        
        Assert.False(result.Item1);
        Assert.NotNull(result.Item2);
        // The error message might be "invalid dimensions" or an exception message
        // depending on how the dynamic binding resolves
        Assert.True(result.Item2.Contains("invalid dimensions") || result.Item2.Contains("Error validating range"));
    }

    [Fact]
    public void TryValidateRange_CustomMaxCells_RespectedLimit()
    {
        dynamic range = MockRange.Create(100, 100); // 10,000 cells
        
        var result = RangeValidator.TryValidateRange(range, maxCells: 5000);
        
        Assert.False(result.Item1);
        Assert.NotNull(result.Item2);
        Assert.Contains("too large", result.Item2);
        Assert.Contains("10,000", result.Item2);
        Assert.Contains("5,000", result.Item2);
        Assert.Equal(100, result.Item3);
        Assert.Equal(100, result.Item4);
        Assert.Equal(10_000, result.Item5);
    }

    #endregion

    #region Security Tests

    [Fact]
    public void ValidateRange_PreventIntegerOverflow_ThrowsForExtremeDimensions()
    {
        // Even though individual dimensions might fit in int, 
        // the multiplication could overflow if not using long
        // This test ensures we handle large dimensions safely
        dynamic range = MockRange.Create(50000, 50000); // Would be 2.5 billion if multiplied
        
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            RangeValidator.ValidateRange(range);
        });
        Assert.Contains("too large", ex.Message);
    }

    #endregion
}
