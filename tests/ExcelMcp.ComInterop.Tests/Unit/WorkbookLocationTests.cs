using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("Layer", "ComInterop")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookLocationTests
{
    [Theory]
    [InlineData("https://contoso.sharepoint.com/sites/Test/Shared Documents/Test $1.xlsx?web=1")]
    [InlineData("https://CONTOSO.sharepoint.com/sites/Test/Shared%20Documents/Test%20%241.xlsx?web=0")]
    [InlineData("https://contoso.sharepoint.com:443/sites/Test/Shared%20Documents/Test%20$1.xlsx")]
    public void Normalize_EquivalentUrls_HaveOneReportedLocation(string input)
    {
        const string expected = "https://contoso.sharepoint.com/sites/Test/Shared%20Documents/Test%20%241.xlsx";
        Assert.Equal(expected, WorkbookLocation.Normalize(input));
        Assert.Equal(expected, WorkbookLocation.Normalize(WorkbookLocation.Normalize(input)));
    }

    [Theory]
    [InlineData("http://example.com/Test.csv", "http://example.com/Test.csv")]
    [InlineData("https://example.com:8443/Test.txt?download=1#fragment", "https://example.com:8443/Test.txt")]
    [InlineData("https://d.docs.live.net/123/Documents/Test $1.xlsx?token=test#fragment",
        "https://d.docs.live.net/123/Documents/Test%20%241.xlsx")]
    [InlineData("https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=test",
        "https://contoso.sharepoint.com/_layouts/15/Doc.aspx")]
    [InlineData("https://contoso.sharepoint.com/:x:/s/Test/sharing-link",
        "https://contoso.sharepoint.com/%3Ax%3A/s/Test/sharing-link")]
    [InlineData("https://example.com/folder%2fTest.xlsx", "https://example.com/folder%2FTest.xlsx")]
    [InlineData("https://example.com/folder%5cTest.xlsx", "https://example.com/folder%5CTest.xlsx")]
    public void Normalize_RemoteLocations_PreserveSanitizedIdentity(string input, string expected)
    {
        Assert.True(WorkbookLocation.IsRemote(input));
        Assert.Equal(expected, WorkbookLocation.Normalize(input));
        Assert.Equal(expected, WorkbookLocation.Normalize(expected));
    }

    [Fact]
    public void Normalize_RemoteIdentity_DoesNotContainCredentialsOrParameters()
    {
        var input = new UriBuilder("https://d.docs.live.net/123/Documents/Test%20File.xlsx")
        {
            UserName = "reader",
            Password = "test-password",
            Query = "token=test-token",
            Fragment = "test-fragment"
        }.Uri.AbsoluteUri;

        Assert.Equal("https://d.docs.live.net/123/Documents/Test%20File.xlsx",
            WorkbookLocation.Normalize(input));
    }

    [Fact]
    public void OwnershipKeys_DistinguishQueriesWithoutExposingTheirValues()
    {
        const string endpoint = "https://example.com/export.xlsx";
        const string first = endpoint + "?report=one&token=test-first";
        const string second = endpoint + "?report=two&token=test-second";
        var firstKey = WorkbookLocation.GetOwnershipKey(first);
        var secondKey = WorkbookLocation.GetOwnershipKey(second);

        Assert.NotEqual(firstKey, secondKey);
        Assert.NotEqual(WorkbookLocation.GetOwnershipKey(endpoint), firstKey);
        Assert.Equal(firstKey, WorkbookLocation.GetOwnershipKey(first + "#worksheet"));
        Assert.Equal(endpoint, WorkbookLocation.Normalize(firstKey));
        Assert.Equal(endpoint, WorkbookLocation.Normalize(secondKey));
        Assert.DoesNotContain("test-first", firstKey);
        Assert.DoesNotContain("report=", firstKey);
        Assert.DoesNotContain("test-second", secondKey);
    }

    [Fact]
    public void OwnershipKeys_LocalAliasesAndFileUrls_UseOneCanonicalPath()
    {
        const string path = @"C:\Books\Test File.xlsx";
        Assert.Equal(path, WorkbookLocation.GetOwnershipKey(@"C:\Books\.\Test File.xlsx"));
        Assert.Equal(path, WorkbookLocation.GetOwnershipKey("file:///C:/Books/Test%20File.xlsx"));
    }

    [Theory]
    [InlineData(@"C:\Books\Test.xlsx", @"C:\Books\Test.xlsx")]
    [InlineData("https://example.com/export.xlsx?report=one&token=test-first", "https://example.com/export.xlsx")]
    [InlineData("https://example.com/export.xlsx?report=two&token=test-second", "https://example.com/export.xlsx")]
    public void OwnershipKeys_CanFindTheSessionAtItsReportedLocation(string openingLocation, string reportedLocation)
    {
        var key = WorkbookLocation.GetOwnershipKey(openingLocation);
        Assert.True(WorkbookLocation.MatchesReportedLocation(key, reportedLocation));
        Assert.False(WorkbookLocation.MatchesReportedLocation(key, reportedLocation + ".other"));
        Assert.False(WorkbookLocation.MatchesReportedLocation(key, null));
    }

    [Theory]
    [InlineData(@"C:\Workbooks\Test.xlsx")]
    [InlineData(@"\\server\share\Test.xlsm")]
    [InlineData(@"C:\Imports\Test.csv")]
    [InlineData(@"C:\Imports\Test%20File.csv")]
    [InlineData(@"C:\Templates\Test.xltx")]
    public void Normalize_WindowsPaths_RemainLocal(string path)
    {
        Assert.Equal(Path.GetFullPath(path), WorkbookLocation.Normalize(path));
        Assert.False(WorkbookLocation.IsRemote(path));
        Assert.Equal(Path.GetExtension(path), WorkbookLocation.GetExtension(path));
    }

    [Fact]
    public void Normalize_WindowsFileUrl_UsesTheLocalFileIdentity()
    {
        Assert.Equal(@"C:\Imports\Test File.csv",
            WorkbookLocation.Normalize("file:///C:/Imports/Test%20File.csv"));
    }

    [Fact]
    public void Normalize_RelativePath_RequiresAnAbsoluteLocation()
    {
        Assert.Throws<ArgumentException>(() => WorkbookLocation.Normalize("relative.xlsx"));
    }
}
