using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
        parameters ??= [];

        EnsureVbaFile(batch.WorkbookPath);

        using var timeoutCts = timeout.HasValue
            ? new CancellationTokenSource(timeout.Value)
            : null;

        Func<OperationResult> execute = () => batch.Execute((ctx, ct) =>
        {
            var originalCulture = CultureInfo.CurrentCulture;
            var originalUiCulture = CultureInfo.CurrentUICulture;
            try
            {
                var excelCulture = CultureInfo.GetCultureInfo("en-US");
                CultureInfo.CurrentCulture = excelCulture;
                CultureInfo.CurrentUICulture = excelCulture;

                // Use late-bound COM dispatch via Type.InvokeMember to avoid dependency on
                // Microsoft.Vbe.Interop.dll, which is not available on Click-to-Run Office
                // installations. The early-bound PIA call ctx.App.Run() triggers assembly
                // resolution of VBE types through the embedded Application interface metadata.
                // Run follows the user's Excel macro settings; it never lowers AutomationSecurity.
                var args = new object[1 + parameters.Length];
                args[0] = procedureName;
                for (int i = 0; i < parameters.Length; i++)
                {
                    args[i + 1] = parameters[i];
                }

                var excelApplicationType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Excel is not installed or not properly registered.");

                excelApplicationType.InvokeMember(
                    "Run",
                    BindingFlags.InvokeMethod | BindingFlags.DoNotWrapExceptions,
                    null,
                    ctx.App,
                    args,
                    excelCulture);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (COMException comEx) when (comEx.ErrorCode == GenericOfficeAutomationError)
            {
                // Excel reports a missing procedure and macros blocked by the user's macro
                // settings with the same HRESULT, so keep Excel's own text and add guidance
                // without claiming which cause applies.
                throw new InvalidOperationException(BuildRunFailureMessage(procedureName, comEx), comEx);
            }
            finally
            {
                CultureInfo.CurrentCulture = originalCulture;
                CultureInfo.CurrentUICulture = originalUiCulture;
            }
        }, timeoutCts?.Token ?? default);

        return ExecuteWithTimeout(execute, timeoutCts);
    }

    internal static string BuildRunFailureMessage(string procedureName, COMException comEx)
    {
        string detail = string.IsNullOrWhiteSpace(comEx.Message) ? "(no description)" : comEx.Message.Trim();
        return
            $"Excel could not run '{procedureName}' (HRESULT 0x{unchecked((uint)comEx.ErrorCode):X8}): {detail} " +
            "If the procedure name is correct and Excel says macros are disabled, Excel blocked macros in this " +
            "workbook because of your Excel macro settings. ExcelMcp follows those settings and never changes them. " +
            "To allow macros for this file, store it in a trusted location or sign its VBA project with a " +
            "certificate you trust, then close this session and open the workbook again. " +
            "VBA code can still be listed, viewed, and edited while macros are blocked.";
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
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? targetComponent = null;
            try
            {
                // PIA gap: VBProject is in Microsoft.Vbe.Interop, not the Excel PIA.
                // No .NET 5+ compatible NuGet package exists for VBE types.
                vbaProject = ((dynamic)ctx.Book).VBProject;
                EnsureProjectUnlocked((object)vbaProject);
                vbComponents = vbaProject.VBComponents;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic? component = null;
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
