using Sbroenne.ExcelMcp.Core.PowerQuery;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

public class PowerQueryIdentityParsingTests
{
    [Theory]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location=A", "a")]
    [InlineData("OLEDB;provider=microsoft.mashup.oledb.1;location=\"A;B\"", "a;b")]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location='A''B'", "a'b")]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location=Bob's Query;Extended Properties=\"\"", "bob's query")]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location=A\0", "a")]
    public void MatchesMashupLocation_ExactCaseInsensitiveLocation_ReturnsTrue(
        string connectionString,
        string queryName)
    {
        Assert.True(PowerQueryHelpers.MatchesMashupLocation(connectionString, queryName));
    }

    [Theory]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location=AA", "A")]
    [InlineData("OLEDB;Provider=Other.Provider;Location=A", "A")]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Other=A", "A")]
    public void MatchesMashupLocation_DifferentIdentity_ReturnsFalse(
        string connectionString,
        string queryName)
    {
        Assert.False(PowerQueryHelpers.MatchesMashupLocation(connectionString, queryName));
    }

    [Fact]
    public void TryReplaceMashupLocation_ExactQuotedLocation_PreservesOtherProperties()
    {
        const string connectionString =
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location=\"A;Old\";Extended Properties=\"Keep;This\"";

        var replaced = PowerQueryHelpers.TryReplaceMashupLocation(
            connectionString,
            "a;old",
            "A;New",
            out var updatedConnectionString);

        Assert.True(replaced);
        Assert.Equal(
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Location=\"A;New\";Extended Properties=\"Keep;This\"",
            updatedConnectionString);
    }
}
