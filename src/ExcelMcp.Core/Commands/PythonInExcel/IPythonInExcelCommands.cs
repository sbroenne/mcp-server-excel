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
/// GET-RESULT reads the computed value back, polling for a short time because cloud execution is not
/// instantaneous (a fresh formula may transiently read as busy/connecting before the result is ready).
/// IMPORTANT: polling is best-effort - Excel exposes no reliable "still computing" signal via COM, so a
/// freshly written formula can briefly read back as a stale default (e.g. 0) that looks like a valid,
/// stable value even though the real cloud result has not arrived yet. If a result looks suspicious
/// (e.g. an unexpected 0/default), call get-result again after a few seconds rather than trusting the
/// first read.
/// DATA BINDING: reference live worksheet data inside the Python code with xl("A1:A6"),
/// xl("Sheet1!A1:A6"), or a named range xl("MyRange") — all work when the formula is authored via this
/// tool, the same as if typed interactively. TIP: xl() returns a pandas DataFrame/Series (not a plain
/// list) unless you pass headers explicitly; prefer pandas methods (.sum()/.mean()/.max()) over Python
/// builtins (sum()/len()) to avoid getting a Series back instead of a scalar total.
/// </summary>
[ServiceCategory("pythoninexcel", "PythonInExcel")]
[McpTool("pythoninexcel", Title = "Python in Excel Operations", Destructive = true, Category = "data",
    Description = "Write and read Microsoft 365 \"Python in Excel\" =PY() formulas. Requires a licensed M365 account with Python in Excel enabled and internet access - Python code executes in Microsoft's cloud sandbox, not locally. SET-FORMULA writes '=PY(code, returnType)' via Range.Formula2 (returnType: 0=Excel Value, 1=Python Object; always pass it explicitly). GET-RESULT reads back the result, polling briefly since cloud execution is not instantaneous; this is best-effort since Excel exposes no reliable \"still computing\" signal via COM, so a freshly written formula may briefly read back as a stale default (e.g. 0) - call get-result again after a few seconds if the value looks suspicious. Reference live worksheet data inside the Python code using xl(\"A1:A6\"), xl(\"Sheet1!A1:A6\"), or a named range xl(\"MyRange\") - this works reliably. TIP: xl() returns a DataFrame/Series, not a plain list, so prefer .sum()/.mean()/.max() methods over Python's builtin sum()/len().")]
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
    /// Reads back the computed result of a =PY() cell, polling briefly since the Python code
    /// executes asynchronously in Microsoft's cloud sandbox and may not be ready immediately
    /// after the formula is (re)written.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Cell address containing the PY() formula</param>
    /// <param name="maxWaitSeconds">Maximum seconds to poll for a non-transient result (default: 15)</param>
    /// <exception cref="InvalidOperationException">If the range cannot be resolved</exception>
    [ServiceAction("get-result")]
    PythonInExcelResult GetResult(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        int maxWaitSeconds = 15);
}
