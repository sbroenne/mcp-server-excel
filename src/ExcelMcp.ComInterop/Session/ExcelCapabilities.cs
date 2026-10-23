using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Excel capabilities cached for one session. Access only on the session's STA thread.
/// </summary>
public sealed class ExcelCapabilities
{
    internal static string AutoSaveWarning(bool? enabled) => enabled switch
    {
        true => "Excel AutoSave is enabled. Excel may save edits automatically, even without an explicit save. Closing without saving cannot undo changes already saved.",
        false => "Excel AutoSave is off.",
        null => "Excel AutoSave status is unavailable. Excel may save edits automatically if AutoSave is enabled."
    };
    private readonly Func<bool> _probeFormula2;
    private bool? _supportsFormula2;

    internal ExcelCapabilities(Func<bool> probeFormula2) => _probeFormula2 = probeFormula2;

    /// <summary>
    /// Whether range formulas support dynamic-array semantics.
    /// Unexpected probe failures propagate and are not cached as lack of support.
    /// </summary>
    public bool SupportsFormula2 => _supportsFormula2 ??= _probeFormula2();

    internal static bool? ReadAutoSaveStatus(Func<bool> read)
    {
        try
        {
            return read();
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
        {
            return null;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The Formula2 getter alone may return Excel's generic error on older Excel, but only after the same cell's legacy Formula getter succeeds. Writes and other probe failures propagate.")]
    internal static bool ProbeFormula2(Func<object?> readFormula, Func<object?> readFormula2)
    {
        // Confirm the probe cell is readable before interpreting Excel's generic error
        // as an unavailable Formula2 getter. This policy must never wrap a write.
        _ = readFormula();
        try
        {
            _ = readFormula2();
            return true;
        }
        catch (COMException ex) when (ex.HResult is
            unchecked((int)0x80020003) or // DISP_E_MEMBERNOTFOUND
            unchecked((int)0x80020006) or // DISP_E_UNKNOWNNAME
            unchecked((int)0x80004001) or // E_NOTIMPL
            unchecked((int)0x800A03EC))   // Older Excel's Formula2 getter
        {
            return false;
        }
    }
}
