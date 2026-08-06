using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Selects Formula2 or the legacy Formula member for one Excel session.
/// </summary>
public static class FormulaCompatibility
{
    private const int DispEMemberNotFound = unchecked((int)0x80020003);
    private const int Excel1004 = unchecked((int)0x800A03EC);
    private static readonly ConditionalWeakTable<IExcelBatch, FormulaMode> Modes = new();

    /// <summary>
    /// Returns whether Formula2 is available after a read-only probe cached for this session.
    /// </summary>
    public static bool SupportsFormula2(IExcelBatch batch, dynamic range) =>
        Modes.GetValue(batch, _ => Probe(range)).UseFormula2;

    /// <summary>
    /// Reads a formula through the mode selected for this session.
    /// </summary>
    public static object Read(IExcelBatch batch, dynamic range) =>
        SupportsFormula2(batch, range) ? range.Formula2 : range.Formula;

    /// <summary>
    /// Writes a formula through the mode selected for this session.
    /// </summary>
    /// <remarks>
    /// Write failures intentionally propagate. Only a read probe may select Formula fallback.
    /// </remarks>
    public static void Write(IExcelBatch batch, dynamic range, object formulas)
    {
        if (SupportsFormula2(batch, range))
        {
            range.Formula2 = formulas;
            return;
        }

        range.Formula = formulas;
    }

    private static FormulaMode Probe(dynamic range)
    {
        try
        {
            _ = range.Formula2;
            return new FormulaMode(useFormula2: true);
        }
        catch (COMException ex) when (ex.HResult is DispEMemberNotFound or Excel1004)
        {
            return new FormulaMode(useFormula2: false);
        }
        catch (RuntimeBinderException)
        {
            return new FormulaMode(useFormula2: false);
        }
    }

    private sealed class FormulaMode(bool useFormula2)
    {
        public bool UseFormula2 { get; } = useFormula2;
    }
}
