using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    /// <inheritdoc/>
    public WorkbookThemeResult GetTheme(IExcelBatch batch) =>
        batch.Execute((ctx, ct) => ReadWorkbookTheme(ctx.Book, batch.WorkbookPath, ct));

    /// <inheritdoc/>
    public WorkbookThemeResult ApplyTheme(IExcelBatch batch, string themePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(themePath);
        if (!Path.IsPathFullyQualified(themePath))
            throw new ArgumentException("themePath must be an absolute path.", nameof(themePath));
        if (!string.Equals(Path.GetExtension(themePath), ".thmx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("themePath must select an Office .thmx theme.", nameof(themePath));
        var path = Path.GetFullPath(themePath);
        if (!File.Exists(path))
            throw new FileNotFoundException("The selected Office theme does not exist.", path);
        return batch.Execute((ctx, ct) =>
        {
            ct.ThrowIfCancellationRequested();
            ctx.Book.ApplyTheme(path);
            return ReadWorkbookTheme(ctx.Book, batch.WorkbookPath, ct);
        });
    }

    private static WorkbookThemeResult ReadWorkbookTheme(Excel.Workbook book, string filePath, CancellationToken ct)
    {
        dynamic? theme = null;
        dynamic? colors = null;
        dynamic? fonts = null;
        try
        {
            // PIA gap: Office-core Theme/ThemeColorScheme/ThemeFontScheme types are not referenced by this project.
            theme = ((dynamic)book).Theme;
            colors = theme.ThemeColorScheme;
            fonts = theme.ThemeFontScheme;
            var result = new WorkbookThemeResult
            {
                Success = true,
                FilePath = filePath
            };
            for (int index = 1; index <= 12; index++)
            {
                ct.ThrowIfCancellationRequested();
                dynamic? color = null;
                try
                {
                    color = colors.Colors(index);
                    result.Colors.Add(new WorkbookThemeColor(index,
                        ((Excel.XlThemeColor)index).ToString(),
                        FormattingHelpers.ColorToHex(Convert.ToInt32(color.RGB))));
                }
                finally
                {
                    ComUtilities.Release(ref color);
                }
            }
            result.MajorFonts = ReadThemeFonts(fonts, true, ct);
            result.MinorFonts = ReadThemeFonts(fonts, false, ct);
            return result;
        }
        finally
        {
            ComUtilities.Release(ref fonts);
            ComUtilities.Release(ref colors);
            ComUtilities.Release(ref theme);
        }
    }

    private static List<WorkbookThemeFont> ReadThemeFonts(dynamic fontScheme, bool major, CancellationToken ct)
    {
        List<WorkbookThemeFont> result = [];
        string[] scripts = ["Latin", "EastAsian", "ComplexScript"];
        for (int index = 1; index <= scripts.Length; index++)
        {
            ct.ThrowIfCancellationRequested();
            dynamic? font = null;
            try
            {
                font = major ? fontScheme.MajorFont(index) : fontScheme.MinorFont(index);
                result.Add(new WorkbookThemeFont(scripts[index - 1], font.Name));
            }
            finally
            {
                ComUtilities.Release(ref font);
            }
        }
        return result;
    }
}
