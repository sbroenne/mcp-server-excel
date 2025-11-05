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
    public async Task<OperationResult> RunAsync(IExcelBatch batch, string procedureName, TimeSpan? timeout, params string[] parameters)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "vba-run"
        };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            return CreateVbaTrustGuidance();
        }

        return await batch.Execute((ctx, ct) =>
        {
            try
            {
                if (parameters.Length == 0)
                {
                    ctx.App.Run(procedureName);
                }
                else
                {
                    object[] paramObjects = parameters.Cast<object>().ToArray();
                    ctx.App.Run(procedureName, paramObjects);
                }

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = batch.WorkbookPath;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error running procedure '{procedureName}': {ex.Message}";
                return result;
            }
        }, timeout: timeout);
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string moduleName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "vba-delete"
        };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            return CreateVbaTrustGuidance();
        }

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? targetComponent = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
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
                    result.Success = false;
                    result.ErrorMessage = $"Module '{moduleName}' not found";
                    return result;
                }

                vbComponents.Remove(targetComponent);

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = batch.WorkbookPath;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting module: {ex.Message}";
                return result;
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
