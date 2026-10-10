using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
    [InlineData("http://contoso.sharepoint.com/sites/Test/Test.xlsx")]
    [InlineData("https://example.com/Test.xlsx")]
    [InlineData("https://contoso.sharepoint.com.example.com/Test.xlsx")]
    [InlineData("https://sharepoint.com/Test.xlsx")]
    [InlineData("https://user:password@contoso.sharepoint.com/Test.xlsx")]
    [InlineData("https://contoso.sharepoint.com:8443/Test.xlsx")]
    [InlineData("https://contoso.sharepoint.com/Test.xlsx#fragment")]
    [InlineData("https://contoso.sharepoint.com/Test.xlsx?download=1")]
    [InlineData("https://contoso.sharepoint.com/Test.xlsx?web=1&token=secret")]
    [InlineData("https://contoso.sharepoint.com/_layouts/15/Doc.aspx")]
    [InlineData("https://contoso.sharepoint.com/:x:/s/Test/sharing-link")]
    [InlineData("https://contoso.sharepoint.com/sites/Test/Shared%20Documents")]
    [InlineData("https://contoso.sharepoint.com/Test.csv")]
    [InlineData("https://contoso.sharepoint.com/folder%2fTest.xlsx")]
    [InlineData("https://contoso.sharepoint.com/folder%5cTest.xlsx")]
    [InlineData("relative.xlsx")]
    public void Normalize_UnsupportedLocations_AreRejected(string input)
    {
        Assert.Throws<ArgumentException>(() => WorkbookLocation.Normalize(input));
    }

    [Theory]
    [InlineData("https://contoso.sharepoint.com/:x:/t/Test/sharing-link?isSPOFile=1")]
    [InlineData("https://contoso.sharepoint.com/:x:/s/Test/sharing-link")]
    [InlineData("https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=opaque-id")]
    [InlineData("https://contoso.sharepoint.com/_layouts/15/Doc.aspx")]
    public void Normalize_UnsupportedLinks_ExplainHowToObtainSupportedLocation(string input)
    {
        var error = Assert.Throws<ArgumentException>(() => WorkbookLocation.Normalize(input));
        Assert.Contains("File > Info > Copy Path", error.Message, StringComparison.Ordinal);
        Assert.Contains("local workbook", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(input, error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("contoso-my.sharepoint.com", ".xlsm")]
    [InlineData("contoso.sharepoint.us", ".xlsb")]
    [InlineData("contoso.sharepoint.de", ".xls")]
    [InlineData("contoso.sharepoint.cn", ".xlsx")]
    [InlineData("contoso.sharepoint-mil.us", ".xlsx")]
    public void Normalize_SharePointClouds_PreserveWorkbookExtension(string host, string extension)
    {
        var url = $"https://{host}/Documents/Test{extension}?web=1";
        var normalized = WorkbookLocation.Normalize(url);
        Assert.Equal($"https://{host}/Documents/Test{extension}", normalized);
        Assert.True(WorkbookLocation.IsRemote(normalized));
        Assert.Equal(extension, WorkbookLocation.GetExtension(normalized));
    }

    [Theory]
    [InlineData(@"C:\Workbooks\Test.xlsx")]
    [InlineData(@"\\server\share\Test.xlsm")]
    public void Normalize_WindowsPaths_RemainLocal(string path)
    {
        Assert.Equal(Path.GetFullPath(path), WorkbookLocation.Normalize(path));
        Assert.False(WorkbookLocation.IsRemote(path));
        Assert.Equal(Path.GetExtension(path), WorkbookLocation.GetExtension(path));
    }

    [Theory]
    [InlineData(".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook)]
    [InlineData(".xlsx", Excel.XlFileFormat.xlOpenXMLStrictWorkbook)]
    [InlineData(".xlsm", Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)]
    [InlineData(".xlsb", Excel.XlFileFormat.xlExcel12)]
    [InlineData(".xls", Excel.XlFileFormat.xlExcel8)]
    [InlineData(".xls", Excel.XlFileFormat.xlExcel9795)]
    [InlineData(".xls", Excel.XlFileFormat.xlWorkbookNormal)]
    public void ValidateOpenedWorkbookFormat_GenuineWorkbooks_AreAccepted(string extension, Excel.XlFileFormat format)
    {
        Assert.Null(Record.Exception(() => WorkbookLocation.ValidateOpenedWorkbookFormat(
            $"https://contoso.sharepoint.com/Documents/Test{extension}", format)));
        Assert.Null(Record.Exception(() => WorkbookLocation.ValidateOpenedWorkbookFormat(
            $@"C:\Workbooks\Test{extension}", format)));
    }

    [Theory]
    [InlineData(Excel.XlFileFormat.xlHtml)]
    [InlineData(Excel.XlFileFormat.xlWebArchive)]
    [InlineData(Excel.XlFileFormat.xlCSV)]
    [InlineData(Excel.XlFileFormat.xlCurrentPlatformText)]
    public void ValidateOpenedWorkbookFormat_WebOrTextResponses_AreRejected(Excel.XlFileFormat format)
    {
        foreach (var location in new[] { @"C:\Workbooks\Test.xlsm", "https://contoso.sharepoint.com/Documents/Test.xlsm" })
        {
            var error = Assert.Throws<InvalidOperationException>(() =>
                WorkbookLocation.ValidateOpenedWorkbookFormat(location, format));
            Assert.Contains("not a supported Excel workbook", error.Message, StringComparison.Ordinal);
            Assert.Contains("No session was created", error.Message, StringComparison.Ordinal);
            Assert.DoesNotContain(location, error.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ValidateOpenedWorkbookFormat_LocalTextImport_RemainsSupported()
    {
        Assert.Null(Record.Exception(() => WorkbookLocation.ValidateOpenedWorkbookFormat(
            @"C:\Imports\Test.csv", Excel.XlFileFormat.xlCSV)));
    }
}
