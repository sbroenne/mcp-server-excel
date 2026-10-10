using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Normalizes workbook identities without treating SharePoint URLs as disk paths.
/// </summary>
internal static class WorkbookLocation
{
    private const string UnsupportedLinkRecovery =
        "Open the file in desktop Excel and use File > Info > Copy Path to obtain its direct workbook URL, " +
        "or supply an existing local workbook path.";

    internal static bool IsRemote(string location) =>
        location.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    internal static string Normalize(string location)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(location);
        if (!Uri.TryCreate(location, UriKind.Absolute, out var uri) || uri.IsFile)
        {
            if (!Path.IsPathFullyQualified(location))
                throw new ArgumentException("Workbook location must be an absolute Windows path or a direct SharePoint HTTPS workbook URL.", nameof(location));
            return Path.GetFullPath(location);
        }

        if (uri.Scheme != Uri.UriSchemeHttps
            || !uri.IsDefaultPort
            || !string.IsNullOrEmpty(uri.UserInfo)
            || !string.IsNullOrEmpty(uri.Fragment)
            || !(uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase)
                || uri.Host.EndsWith(".sharepoint.us", StringComparison.OrdinalIgnoreCase)
                || uri.Host.EndsWith(".sharepoint.de", StringComparison.OrdinalIgnoreCase)
                || uri.Host.EndsWith(".sharepoint.cn", StringComparison.OrdinalIgnoreCase)
                || uri.Host.EndsWith(".sharepoint-mil.us", StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException("Use a direct SharePoint or OneDrive for Business HTTPS workbook URL without credentials, fragments, or a custom port.", nameof(location));
        }

        if (uri.Query.Length > 0 && uri.Query is not ("?web=1" or "?web=0"))
            throw new ArgumentException(
                "Sharing links and URL parameters other than web=1 or web=0 are not supported. " + UnsupportedLinkRecovery,
                nameof(location));

        var segments = uri.AbsolutePath.Split('/');
        for (var i = 0; i < segments.Length; i++)
        {
            var segment = Uri.UnescapeDataString(segments[i]);
            if (segment.Contains('/') || segment.Contains('\\'))
                throw new ArgumentException("Encoded path separators are not supported in SharePoint workbook URLs.", nameof(location));
            segments[i] = Uri.EscapeDataString(segment);
        }

        var extension = Path.GetExtension(Uri.UnescapeDataString(segments[^1])).ToLowerInvariant();
        if (extension is not (".xlsx" or ".xlsm" or ".xlsb" or ".xls"))
            throw new ArgumentException(
                "Supply a direct SharePoint workbook URL ending in .xlsx, .xlsm, .xlsb or .xls, not a folder or browser/sharing page. " +
                UnsupportedLinkRecovery, nameof(location));

        return new UriBuilder(uri) { Path = string.Join('/', segments), Query = string.Empty }.Uri.AbsoluteUri;
    }

    internal static string GetExtension(string location) =>
        Path.GetExtension(IsRemote(location)
            ? Uri.UnescapeDataString(new Uri(location).AbsolutePath)
            : location).ToLowerInvariant();

    internal static void ValidateOpenedWorkbookFormat(string location, Excel.XlFileFormat format)
    {
        if (GetExtension(location) is not (".xlsx" or ".xlsm" or ".xlsb" or ".xls"))
            return;

        if (format is Excel.XlFileFormat.xlOpenXMLWorkbook
            or Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
            or Excel.XlFileFormat.xlOpenXMLStrictWorkbook
            or Excel.XlFileFormat.xlExcel12
            or Excel.XlFileFormat.xlExcel8
            or Excel.XlFileFormat.xlExcel9795
            or Excel.XlFileFormat.xlExcel5
            or Excel.XlFileFormat.xlExcel4Workbook
            or Excel.XlFileFormat.xlExcel4
            or Excel.XlFileFormat.xlExcel3
            or Excel.XlFileFormat.xlExcel2
            or Excel.XlFileFormat.xlWorkbookNormal)
            return;

        throw new InvalidOperationException(
            "Excel opened a file that is not a supported Excel workbook. " +
            "A sign-in page, browser page, or text response may have opened instead. " +
            "No session was created. " + UnsupportedLinkRecovery);
    }
}
