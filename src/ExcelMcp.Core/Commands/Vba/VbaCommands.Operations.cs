using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script operations (Run)
/// </summary>
public partial class VbaCommands
{
    /// <inheritdoc />
    public OperationResult Run(IExcelBatch batch, string procedureName, TimeSpan? timeout, params string[] parameters)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(procedureName);
        if (procedureName.Contains('!'))
        {
            throw new ArgumentException(
                "Use an unqualified procedure name such as Module1.MySub. The selected session determines the workbook.",
                nameof(procedureName));
        }
        parameters ??= [];
        if (parameters.Length > 30)
        {
            throw new ArgumentException("Excel accepts at most 30 macro parameters.", nameof(parameters));
        }

        EnsureVbaFile(batch.WorkbookPath);

        using var timeoutCts = timeout.HasValue
            ? new CancellationTokenSource(timeout.Value)
            : null;

        Func<OperationResult> execute = () => batch.Execute((ctx, ct) =>
        {
            var owner = batch is AttachedWorkbookBatch attached ? attached.Owner : batch as ExcelBatch;
            if (owner?.RequiresVbaApproval(ctx.Book) == true)
                throw new InvalidOperationException(
                    batch is AttachedWorkbookBatch
                    ? "VBA macros were disabled when this workbook opened. Use Excel to review macro permissions " +
                    "and run the procedure. This workbook was already open in Excel, so this session cannot reopen it " +
                    "to enable macros. No procedure was run."
                    : "VBA macros were disabled when this workbook opened. Obtain approval for file enable-macros " +
                    "with approved:true and allow_reopen:true (CLI: session enable-macros --approved --allow-reopen) " +
                    "before running a procedure. The saved workbook must be reopened; unsaved changes are not discarded.");
            var originalCulture = CultureInfo.CurrentCulture;
            var originalUiCulture = CultureInfo.CurrentUICulture;
            try
            {
                var excelCulture = CultureInfo.GetCultureInfo("en-US");
                CultureInfo.CurrentCulture = excelCulture;
                CultureInfo.CurrentUICulture = excelCulture;

                string qualifiedName = $"'{ctx.Book.Name.Replace("'", "''", StringComparison.Ordinal)}'!{procedureName}";
                var args = new object[30];
                Array.Fill(args, Type.Missing);
                for (int i = 0; i < parameters.Length; i++)
                {
                    args[i] = parameters[i];
                }

                ctx.App.Run(qualifiedName,
                    args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9],
                    args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19],
                    args[20], args[21], args[22], args[23], args[24], args[25], args[26], args[27], args[28], args[29]);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                CultureInfo.CurrentCulture = originalCulture;
                CultureInfo.CurrentUICulture = originalUiCulture;
            }
        }, timeoutCts?.Token ?? default);

        return ExecuteWithTimeout(execute, timeoutCts);
    }

    private static OperationResult ExecuteWithTimeout(
        Func<OperationResult> operation,
        CancellationTokenSource? timeoutCts)
    {
        try
        {
            return operation();
        }
        catch (OperationCanceledException ex) when (timeoutCts?.IsCancellationRequested == true)
        {
            throw new TimeoutException("VBA procedure execution exceeded its requested timeout.", ex);
        }
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string moduleName)
    {
        EnsureVbaFile(batch.WorkbookPath);

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? vbaProject = null;
            Vbe.VBComponents? vbComponents = null;
            Vbe.VBComponent? targetComponent = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
                EnsureProjectUnlocked(vbaProject);
                vbComponents = vbaProject.VBComponents;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    Vbe.VBComponent? component = null;
                    try
                    {
                        component = vbComponents.Item(i);
                        if (component.Name == moduleName)
                        {
                            targetComponent = component;
                            component = null; // Don't release - we're keeping it
                            break;
                        }
                    }
                    finally
                    {
                        if (component != null)
                        {
                            ComUtilities.Release(ref component);
                        }
                    }
                }

                if (targetComponent == null)
                {
                    throw new OperationFailureException(OperationFailureCategory.NotFound, $"Module '{moduleName}' not found.");
                }

                vbComponents.Remove(targetComponent);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (COMException comEx) when (IsVbaTrustError(comEx))
            {
                throw new OperationFailureException(OperationFailureCategory.Permissions, VbaTrustErrorMessage, comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                // Trust passed IsVbaTrustError, so this generic 0x800A03EC is a real,
                // non-trust failure. Surface it with COM environment diagnostics (issue #671).
                throw new InvalidOperationException(BuildGenericComErrorMessage(comEx), comEx);
            }
            finally
            {
                ComUtilities.Release(ref targetComponent);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }
}
