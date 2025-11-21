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
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "vba-run"
        };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            throw new ArgumentException(validationError, nameof(batch));
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            return CreateVbaTrustGuidance();
        }

        return batch.Execute((ctx, ct) =>
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
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string moduleName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "vba-delete"
        };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            throw new InvalidOperationException(validationError);
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            throw new InvalidOperationException(trustGuidance.ErrorMessage);
        }

        return batch.Execute((ctx, ct) =>
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
                    throw new InvalidOperationException($"Module '{moduleName}' not found");
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
            finally
            {
                ComUtilities.Release(ref targetComponent);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }
}

