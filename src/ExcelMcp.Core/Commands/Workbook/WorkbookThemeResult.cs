using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Native workbook theme; theme-sensitive cells and styles resolve against these colors and fonts.</summary>
public sealed class WorkbookThemeResult : OperationResult
{
    /// <summary>Excel COM does not expose the applied theme's file/name or scheme names.</summary>
    public string ReadLimitations { get; set; } = "Excel COM does not expose the applied theme file/name or scheme names; returned colors and fonts are native definitions.";
    /// <summary>Every native theme color, indices 1 through 12.</summary>
    public List<WorkbookThemeColor> Colors { get; set; } = [];
    /// <summary>Major Latin, East Asian, and complex-script font definitions.</summary>
    public List<WorkbookThemeFont> MajorFonts { get; set; } = [];
    /// <summary>Minor Latin, East Asian, and complex-script font definitions.</summary>
    public List<WorkbookThemeFont> MinorFonts { get; set; } = [];
}

/// <summary>One native color scheme slot.</summary>
/// <param name="Index">Native theme color index.</param>
/// <param name="Name">Native Excel theme color name.</param>
/// <param name="Rgb">Resolved #RRGGBB color.</param>
public sealed record WorkbookThemeColor(int Index, string Name, string Rgb);

/// <summary>One native font scheme script; empty names are native unspecified definitions, not guessed fallback fonts.</summary>
/// <param name="Script">Latin, EastAsian, or ComplexScript.</param>
/// <param name="Name">Native font name, possibly empty.</param>
public sealed record WorkbookThemeFont(string Script, string Name);
