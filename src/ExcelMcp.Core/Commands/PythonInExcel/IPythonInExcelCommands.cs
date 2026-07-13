using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;

/// <summary>
/// Microsoft 365 "Python in Excel" (=PY()) formulas.
/// REQUIRES: a real Excel session signed into a licensed Microsoft 365 account with Python in Excel
/// enabled, plus internet access — the Python code executes in a Microsoft-hosted cloud sandbox, not
/// locally. Not available offline or with perpetual-license Excel.
/// SET-FORMULA writes '=PY("&lt;code&gt;", returnType)' via Range.Formula2. returnType 0 = "Excel Value"
/// (a plain value/array), 1 = "Python Object" (a rich data type card, e.g. a DataFrame). The returnType
/// argument must always be passed explicitly — omitting it causes a #NAME? error.
/// GET-RESULT reads the computed value back, polling until the cloud round-trip finishes (a fresh
/// formula reads as #BUSY! while the cloud Python sandbox is still computing). Completion is detected
/// deterministically from Excel's calculation state plus a per-cell #BUSY! guard, so a converged
/// result is not confused with the in-flight #BUSY! placeholder. If the cloud backend is still busy
/// at the deadline (e.g. a cold start), GET-RESULT reports that and asks the caller to retry or raise
/// maxWaitSeconds rather than returning a placeholder.
/// DATA BINDING: reference live worksheet data inside the Python code with xl("A1:A6"),
/// xl("Sheet1!A1:A6"), or a named range xl("MyRange") — all work when the formula is authored via this
/// tool, the same as if typed interactively. TIP: xl() returns a pandas DataFrame/Series (not a plain
/// list) unless you pass headers explicitly; prefer pandas methods (.sum()/.mean()/.max()) over Python
/// builtins (sum()/len()) to avoid getting a Series back instead of a scalar total.
/// </summary>
[ServiceCategory("pythoninexcel", "PythonInExcel")]
[McpTool("pythoninexcel", Title = "Python in Excel Operations", Destructive = true, Category = "data",
    Description = "Write and read Microsoft 365 \"Python in Excel\" =PY() formulas. Requires a licensed M365 account with Python in Excel enabled and internet access - Python code executes in Microsoft's cloud sandbox, not locally. SET-FORMULA writes '=PY(code, returnType)' via Range.Formula2 (returnType: 0=Excel Value, 1=Python Object; always pass it explicitly). GET-RESULT reads back the result, polling until the cloud round-trip completes (a fresh formula reads as #BUSY! while still computing); completion is detected deterministically from Excel's calculation state, so a real result is not confused with the #BUSY! placeholder. If the backend is still busy at the deadline (e.g. a cold start), GET-RESULT says so - call it again or raise maxWaitSeconds. Reference live worksheet data inside the Python code using xl(\"A1:A6\"), xl(\"Sheet1!A1:A6\"), or a named range xl(\"MyRange\") - this works reliably. TIP: xl() returns a DataFrame/Series, not a plain list, so prefer .sum()/.mean()/.max() methods over Python's builtin sum()/len().")]
public interface IPythonInExcelCommands
{
    /// <summary>
    /// Writes a =PY() formula to a cell/range via Range.Formula2. The Python code string is
    /// automatically quote-escaped for embedding inside the Excel formula.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Target cell address (e.g. "D1")</param>
    /// <param name="code">Python source code (e.g. "xl('A1:A6').sum()")</param>
    /// <param name="returnType">0 = Excel Value (default), 1 = Python Object</param>
    /// <exception cref="InvalidOperationException">If the range cannot be resolved</exception>
    [ServiceAction("set-formula")]
    OperationResult SetFormula(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter] string code,
        int returnType = 0);

    /// <summary>
    /// Reads back the computed result of a =PY() cell, polling until the Python code (which executes
    /// asynchronously in Microsoft's cloud sandbox) finishes. Completion is detected deterministically
    /// from Excel's calculation state plus a per-cell #BUSY! guard, so a converged result is never
    /// confused with the in-flight #BUSY! placeholder. If the cloud backend is still busy at the
    /// deadline, the result reports failure and asks the caller to retry or raise maxWaitSeconds.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Cell address containing the PY() formula</param>
    /// <param name="maxWaitSeconds">Maximum seconds to poll for the cloud result before giving up (default: 30). Returns as soon as the result is ready.</param>
    /// <exception cref="InvalidOperationException">If the range cannot be resolved</exception>
    [ServiceAction("get-result")]
    PythonInExcelResult GetResult(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        int maxWaitSeconds = 30);
}
