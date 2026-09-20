using System.Globalization;
using Xunit;

namespace Sbroenne.ExcelMcp.Tests.Helpers;

/// <summary>
/// Skips a locale-specific integration test unless the host runs under the ja-JP culture.
/// </summary>
public sealed class JapaneseLocaleFactAttribute : FactAttribute
{
    private const string UnsupportedLocaleMessage =
        "Run this regression on a ja-JP Windows and Excel installation to exercise the locale-specific round trip.";

    /// <summary>
    /// Initializes a new instance of the <see cref="JapaneseLocaleFactAttribute"/> class.
    /// </summary>
    public JapaneseLocaleFactAttribute()
    {
        if (!CultureInfo.CurrentCulture.Name.Equals("ja-JP", StringComparison.OrdinalIgnoreCase))
        {
            Skip = UnsupportedLocaleMessage;
        }
    }
}
