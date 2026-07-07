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
    public PythonInExcelResult GetResult(IExcelBatch batch, string sheetName, string rangeAddress, int maxWaitSeconds = 15)
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

                // The Python code executes in a Microsoft-hosted cloud sandbox, not locally. There is
                // no reliable "still computing" signal exposed via COM: Excel's local calculation
                // engine considers itself done as soon as it dispatches the async call, so the cell
                // shows a stale/default placeholder (e.g. plain "0", or even a transient error-shaped
                // Value2) before the real cloud result ever arrives - and that placeholder can remain
                // "stable" (identical across repeated reads) for a long time, indistinguishable from a
                // genuinely converged result using only Value2/Text. A minimum settle window plus a
                // longer stable-sample streak reduce (but cannot fully eliminate) false positives - this
                // is a best-effort heuristic, not a guarantee. Classification below only trusts the
                // final read once `stable` is true; if the deadline is reached without ever observing
                // the required run of matching reads, GetResult reports failure and asks the caller to
                // retry rather than guessing at an unconverged value.
                // NOTE: Application.Wait() was tried as a message-pumping alternative to Thread.Sleep()
                // (in case the cloud callback needs Excel's message queue pumped to be delivered) but it
                // hung indefinitely in this non-visible automation context, so Thread.Sleep() is used.
                const int MinSettleSeconds = 10;
                const int RequiredStableSamples = 4;
                const int PollIntervalMs = 2000;

                // Well-known Excel error code for #PYTHON! (Python code raised an error), used only to
                // look up the canonical human-readable message via MapErrorCodeToMessage. Detection of
                // "this is a Python error" is NOT done by matching this exact code (the actual negative
                // int Value2 returns for a Python-side error/object has been observed to vary and is not
                // a reliable discriminator on its own) - see the returnType-based classification below.
                const int PythonErrorCode = -2146826233;

                object? value = null;
                string text = string.Empty;
                object? previousValue = null;
                int consecutiveMatches = 0;
                bool stable = false;

                var startTime = DateTime.UtcNow;
                var deadline = startTime.AddSeconds(Math.Max(MinSettleSeconds + 1, maxWaitSeconds));
                do
                {
                    ctx.App.CalculateFullRebuild();
                    value = range.Value2;
                    text = range.Text?.ToString() ?? string.Empty;

                    bool isTransient = Array.IndexOf(TransientMarkers, text) >= 0;
                    // Match on Value2 only - range.Text has been observed to render inconsistently
                    // (e.g. sometimes blank, sometimes a formatted error string) even while Value2
                    // itself has already converged to a stable value/error code.
                    bool matchesPrevious = previousValue != null && Equals(previousValue, value);

                    consecutiveMatches = isTransient ? 0 : matchesPrevious ? consecutiveMatches + 1 : 1;

                    bool minTimeElapsed = (DateTime.UtcNow - startTime).TotalSeconds >= MinSettleSeconds;
                    if (!isTransient && consecutiveMatches >= RequiredStableSamples && minTimeElapsed)
                    {
                        stable = true;
                        break;
                    }

                    previousValue = value;
                    Thread.Sleep(PollIntervalMs);
                }
                while (DateTime.UtcNow < deadline);

                if (!stable)
                {
                    // Never observed the required run of matching reads - the last read cannot be
                    // trusted as the converged Python result (it may be a stale/default placeholder,
                    // or an in-flight transient error code). Report failure rather than guessing.
                    result.Success = false;
                    result.ErrorMessage = $"Python in Excel result did not stabilize within {maxWaitSeconds}s and may not reflect the completed cloud computation. Try get-result again, or increase maxWaitSeconds.";
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
