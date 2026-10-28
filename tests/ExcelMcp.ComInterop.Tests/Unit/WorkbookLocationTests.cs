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
    public void Normalize_EquivalentUrls_HaveOneSessionIdentity(string input)
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
