using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
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
        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            throw new ArgumentException(validationError, nameof(batch));
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            throw new InvalidOperationException(VbaTrustErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
        {
            try
            {
                // Use late-bound COM dispatch (IDispatch) to avoid loading
                // Microsoft.Vbe.Interop.dll, which is not available on Office 365
                // Click-to-Run installations. This mirrors the late-binding approach
                // used by all other VBA operations (List, View, Import, Update, Delete)
                // that access ctx.Book.VBProject via ((dynamic)ctx.Book).
                if (parameters.Length == 0)
                {
                    ((dynamic)ctx.App).Run(procedureName);
                }
                else
                {
                    // Application.Run(MacroName, Arg1, ..., Arg30): build the full
                    // argument array and dispatch via InvokeMember so a variable-length
                    // parameter list is handled correctly.  This also fixes the previous
                    // implementation that incorrectly passed the parameter array as a
                    // single second argument instead of spreading individual parameters.
                    object[] runArgs = new object[parameters.Length + 1];
                    runArgs[0] = procedureName;
                    for (int i = 0; i < parameters.Length; i++)
                    {
                        runArgs[i + 1] = parameters[i];
                    }

                    ctx.App.GetType().InvokeMember(
                        "Run",
                        BindingFlags.InvokeMethod,
                        null,
                        ctx.App,
                        runArgs,
                        CultureInfo.InvariantCulture);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (TargetInvocationException tie)
                when (tie.InnerException is COMException comEx &&
                      (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                       comEx.ErrorCode == unchecked((int)0x800A03EC)))
            {
                throw new InvalidOperationException(VbaTrustErrorMessage, tie.InnerException);
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                // Re-throw the original COM exception from InvokeMember, preserving
                // the original stack trace and exception type.
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
                throw; // unreachable - satisfies the compiler
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                throw new InvalidOperationException(VbaTrustErrorMessage, comEx);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string moduleName)
    {
        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            throw new InvalidOperationException(validationError);
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            throw new InvalidOperationException(VbaTrustErrorMessage);
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
                    throw new InvalidOperationException($"Module '{moduleName}' not found.");
                }

                vbComponents.Remove(targetComponent);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                throw new InvalidOperationException(VbaTrustErrorMessage, comEx);
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



