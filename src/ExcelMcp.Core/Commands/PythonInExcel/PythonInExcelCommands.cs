using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;

/// <summary>
/// Implementation of Microsoft 365 "Python in Excel" (=PY()) formula commands.
/// </summary>
public sealed class PythonInExcelCommands : IPythonInExcelCommands
{
    /// <summary>
    /// Transient result markers returned by Excel while the cloud Python sandbox is still
    /// computing/connecting - not real errors, just "not ready yet".
    /// </summary>
    private static readonly string[] TransientMarkers = ["#BUSY!", "#CONNECT!", "#BLOCKED!"];

    /// <inheritdoc />
    public OperationResult SetFormula(IExcelBatch batch, string sheetName, string rangeAddress, string code, int returnType = 0)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            throw new ArgumentException("Python code must not be empty.", nameof(code));
        }

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-formula" };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Escape embedded double quotes for the Excel formula string literal (" -> "")
                // The returnType argument (0=Excel Value, 1=Python Object) must always be passed
                // explicitly - omitting it causes a #NAME? error when the formula is set via COM.
                string escapedCode = code.Replace("\"", "\"\"", StringComparison.Ordinal);
                string formula = $"=PY(\"{escapedCode}\",{returnType})";

                range.Formula2 = formula;

                result.Success = true;
                result.Message = $"Set Python in Excel formula on '{range.Address}'. Use get-result to read the computed value once the cloud Python backend finishes.";
                return result;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot set Python in Excel formula on range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public PythonInExcelResult GetResult(IExcelBatch batch, string sheetName, string rangeAddress, int maxWaitSeconds = 30)
    {
        var result = new PythonInExcelResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                result.RangeAddress = range.Address;

                string formula = range.Formula2?.ToString() ?? string.Empty;
                result.Formula = formula;

                if (!formula.Contains("PY(", StringComparison.Ordinal))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Cell '{result.RangeAddress}' does not contain a Python in Excel (PY()) formula.";
                    return result;
                }

                // Completion detection uses Excel's calculation state plus a per-cell #BUSY! guard -
                // NOT a value-stability guess.
                //
                // Ground truth (verified empirically against current Excel, both visible and headless):
                // while the Microsoft-hosted cloud Python sandbox is still computing, the cell reads back
                // as the #BUSY! placeholder (Range.Value2 == -2146826237) and Application.CalculationState
                // is xlPending/xlCalculating. The moment the real result arrives, Value2 flips to the
                // computed value (or an error code such as #PYTHON!) and CalculationState returns to
                // xlDone. So the cell is "done" precisely when CalculationState == xlDone AND it no longer
                // reads #BUSY!. This is deterministic and does not lock onto a stale placeholder the way
                // the previous value-stability heuristic could.
                //
                // We deliberately do NOT call CalculateFullRebuild() on every poll: re-dispatching the
                // async PY call repeatedly can keep the cell perpetually #BUSY! and prevent convergence.
                // Instead we nudge calculation once (in case the workbook is in manual-calc mode) and then
                // just observe. Thread.Sleep() is used to wait between reads - Application.Wait() was tried
                // previously but hung indefinitely in the non-visible automation context.

                // #BUSY! - the cloud Python call is still in flight (not a real error).
                const int BusyErrorCode = -2146826237;
                // XlCalculationState.xlDone (calculation complete).
                const int XlDone = 0;
                const int PollIntervalMs = 500;
                // Consecutive non-busy reads that let us trust the result even if the workbook's
                // application-level CalculationState is held pending by OTHER async cells (e.g. a Power
                // Query or a second PY cell). This keeps convergence dependent primarily on THIS cell.
                const int RequiredNonBusyReads = 3;

                // Well-known Excel error code for #PYTHON! (Python code raised an error), used only to
                // look up the canonical human-readable message via MapErrorCodeToMessage. Detection of
                // "this is a Python error" is NOT done by matching this exact code (the actual negative
                // int Value2 returns for a Python-side error/object has been observed to vary and is not
                // a reliable discriminator on its own) - see the returnType-based classification below.
                const int PythonErrorCode = -2146826233;

                // Nudge calculation once so a manual-calc workbook actually dispatches the async PY call.
                // Harmless if calculation is automatic or already running.
                try
                {
                    ctx.App.Calculate();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Application is busy calculating - the nudge is best-effort, so ignore and poll.
                }

                object? value = null;
                string text = string.Empty;
                int nonBusyReads = 0;
                bool converged = false;
                // Last state observed by the poll loop, used to build an accurate timeout diagnostic
                // instead of hard-coding a "#BUSY!" assumption (the cell can also stall on #CONNECT! or
                // #BLOCKED!, or read a settled value while the application is still calculating).
                int calcState = XlDone;
                string lastMarker = string.Empty;

                var deadline = DateTime.UtcNow.AddSeconds(Math.Max(1, maxWaitSeconds));
                do
                {
                    value = range.Value2;
                    text = range.Text?.ToString() ?? string.Empty;
                    calcState = (int)ctx.App.CalculationState;

                    int markerIndex = Array.IndexOf(TransientMarkers, text);
                    bool cellBusy = (value is int busyCode && busyCode == BusyErrorCode)
                        || markerIndex >= 0;
                    lastMarker = value is int c && c == BusyErrorCode ? "#BUSY!"
                        : markerIndex >= 0 ? TransientMarkers[markerIndex]
                        : string.Empty;

                    nonBusyReads = cellBusy ? 0 : nonBusyReads + 1;

                    // Done when this cell is no longer #BUSY! and either the application has finished
                    // calculating or the cell has read a settled (non-busy) value across several
                    // consecutive polls (guards against unrelated pending async cells).
                    if (!cellBusy && (calcState == XlDone || nonBusyReads >= RequiredNonBusyReads))
                    {
                        converged = true;
                        break;
                    }

                    Thread.Sleep(PollIntervalMs);
                }
                while (DateTime.UtcNow < deadline);

                if (!converged)
                {
                    // Deadline reached without a settled result. Report exactly what was last observed
                    // (the transient marker if any, plus the application calculation state) so callers
                    // can distinguish a cold-start #BUSY! stall from a #CONNECT!/#BLOCKED! condition or a
                    // still-pending calculation, rather than always blaming #BUSY!.
                    string observed = lastMarker.Length > 0
                        ? $"the cell still reads as {lastMarker}"
                        : calcState != XlDone
                            ? "the workbook is still calculating"
                            : "the result did not settle";
                    result.Success = false;
                    result.ErrorMessage = $"Python in Excel result did not finish within {maxWaitSeconds}s ({observed}). The Microsoft-hosted Python backend may be under cold-start load - call get-result again, or increase maxWaitSeconds.";
                    return result;
                }

                // Which =PY(code, returnType) mode was requested - parsed from the formula text itself
                // rather than trusting range.Text (which has been observed to render inconsistently in
                // non-visible automation, seemingly because it depends on screen repaint rather than
                // the underlying calculated value). 0 = Excel Value, 1 = Python Object.
                var returnTypeMatch = System.Text.RegularExpressions.Regex.Match(formula, @",\s*(\d+)\s*\)\s*$");
                int formulaReturnType = returnTypeMatch.Success
                    ? int.Parse(returnTypeMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)
                    : 0;

                if (value is int errorCode && errorCode < 0)
                {
                    // Standard Excel error codes are well-known/fixed and mean the *formula itself*
                    // failed to evaluate (e.g. bad range reference) - these are never Python results.
                    bool isStandardExcelError = errorCode is -2146826288 or -2147483648 or -2146826259
                        or -2146826246 or -2146826252 or -2142019887;

                    if (isStandardExcelError)
                    {
                        result.Success = false;
                        result.ErrorMessage = RangeCommands.MapErrorCodeToMessage(errorCode);
                    }
                    else if (formulaReturnType == 1)
                    {
                        // Python Object mode: Value2 is always an error-shaped placeholder for rich
                        // data types (e.g. a DataFrame) since COM cannot represent them. Text would
                        // normally carry a type-name label (e.g. "DataFrame") but is not reliable
                        // enough here to depend on - report the object without assuming its type.
                        result.Success = true;
                        result.IsPythonObject = true;
                        result.TypeName = !string.IsNullOrEmpty(text) && text.Any(char.IsLetter) ? text : null;
                        result.Message = "Cell holds a Python Object (rich data type such as a DataFrame). "
                            + "Value2 cannot expose rich Python object data via COM automation - set returnType=0 "
                            + "(Excel Value) instead if you need to read the underlying data.";
                    }
                    else
                    {
                        // Excel Value mode (returnType=0) with a non-standard negative error code -
                        // this is the Python code itself raising an error (syntax or runtime exception).
                        result.Success = false;
                        result.IsPythonError = true;
                        result.ErrorMessage = RangeCommands.MapErrorCodeToMessage(PythonErrorCode);
                    }
                }
                else
                {
                    result.Success = true;
                    result.Value = value;
                }

                return result;
            }
            catch (System.Runtime.InteropServices.COMException comEx) when (comEx.HResult == unchecked((int)0x8007000E))
            {
                // E_OUTOFMEMORY - Excel's misleading error for sheet/range/session issues
                throw new InvalidOperationException($"Cannot read Python in Excel result from range '{rangeAddress}' on sheet '{sheetName}': {comEx.Message}", comEx);
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }
}
