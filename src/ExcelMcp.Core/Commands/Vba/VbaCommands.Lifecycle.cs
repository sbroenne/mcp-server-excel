using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script lifecycle operations (List, View, Export, Import, Update, Delete)
/// </summary>
public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaListResult List(IExcelBatch batch)
    {
        var result = new VbaListResult { FilePath = batch.WorkbookPath };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            // For LLM-friendly behavior: .xlsx files don't support VBA, return empty list instead of error
            result.Success = true;
            result.Scripts = [];
            return result;
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
            try
            {
                // PIA gap: VBProject is in Microsoft.Vbe.Interop, not the Excel PIA.
                // No .NET 5+ compatible NuGet package exists for VBE types (ThammimTech.Microsoft.Vbe.Interop targets .NET Framework only).
                vbaProject = ((dynamic)ctx.Book).VBProject;
                vbComponents = vbaProject.VBComponents;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic? component = null;
                    dynamic? codeModule = null;
                    try
                    {
                        component = vbComponents.Item(i);
                        string name = component.Name;
                        int type = component.Type;

                        string typeStr = GetVbaModuleTypeName(type);

                        var procedures = new List<string>();
                        codeModule = component.CodeModule;
                        int moduleLineCount = codeModule.CountOfLines;

                        // Parse procedures from code
                        for (int line = 1; line <= moduleLineCount; line++)
                        {
                            string codeLine = codeModule.Lines[line, 1];
                            string trimmedLine = codeLine.TrimStart();
                            if (trimmedLine.StartsWith("Sub ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Function ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Public Sub ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Public Function ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Private Sub ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Private Function ", StringComparison.Ordinal))
                            {
                                string procName = ExtractProcedureName(codeLine);
                                if (!string.IsNullOrEmpty(procName))
                                {
                                    procedures.Add(procName);
                                }
                            }
                        }

                        result.Scripts.Add(new ScriptInfo
                        {
                            Name = name,
                            Type = typeStr,
                            LineCount = moduleLineCount,
                            Procedures = procedures
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref codeModule);
                        ComUtilities.Release(ref component);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                throw new InvalidOperationException(VbaTrustErrorMessage, comEx);
            }
            finally
            {
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public VbaViewResult View(IExcelBatch batch, string moduleName)
    {
        var result = new VbaViewResult { FilePath = batch.WorkbookPath, ModuleName = moduleName };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            throw new InvalidOperationException(validationError);
        }

        if (string.IsNullOrWhiteSpace(moduleName))
        {
            throw new ArgumentException("Module name cannot be empty", nameof(moduleName));
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
            dynamic? component = null;
            dynamic? codeModule = null;
            try
            {
                // PIA gap: VBProject is in Microsoft.Vbe.Interop, not the Excel PIA.
                // No .NET 5+ compatible NuGet package exists for VBE types.
                vbaProject = ((dynamic)ctx.Book).VBProject;
                vbComponents = vbaProject.VBComponents;

                // Find the specified module
                bool found = false;
                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    component = vbComponents.Item(i);
                    string name = component.Name;

                    if (name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        int type = component.Type;
                        result.ModuleType = GetVbaModuleTypeName(type);

                        codeModule = component.CodeModule;
                        result.LineCount = codeModule.CountOfLines;

                        // Read all code lines
                        if (result.LineCount > 0)
                        {
                            result.Code = codeModule.Lines[1, result.LineCount];
                        }

                        // Parse procedures
                        for (int line = 1; line <= result.LineCount; line++)
                        {
                            string codeLine = codeModule.Lines[line, 1];
                            string trimmedLine = codeLine.TrimStart();
                            if (trimmedLine.StartsWith("Sub ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Function ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Public Sub ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Public Function ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Private Sub ", StringComparison.Ordinal) ||
                                trimmedLine.StartsWith("Private Function ", StringComparison.Ordinal))
                            {
                                string procName = ExtractProcedureName(codeLine);
                                if (!string.IsNullOrEmpty(procName))
                                {
                                    result.Procedures.Add(procName);
                                }
                            }
                        }

                        break;
                    }

                    ComUtilities.Release(ref component);
                    component = null;
                }

                if (!found)
                {
                    throw new InvalidOperationException($"Module '{moduleName}' not found in workbook");
                }

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                throw new InvalidOperationException(VbaTrustErrorMessage, comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Import(IExcelBatch batch, string moduleName, string vbaCode)
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
            dynamic? newModule = null;
            dynamic? codeModule = null;
            try
            {
                // PIA gap: VBProject is in Microsoft.Vbe.Interop, not the Excel PIA.
                // No .NET 5+ compatible NuGet package exists for VBE types.
                vbaProject = ((dynamic)ctx.Book).VBProject;
                vbComponents = vbaProject.VBComponents;

                // Check if module already exists
                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic? component = null;
                    try
                    {
                        component = vbComponents.Item(i);
                        if (component.Name == moduleName)
                        {
                            throw new InvalidOperationException($"Module '{moduleName}' already exists. Use script-update to modify it.");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref component);
                    }
                }

                // Add new module
                newModule = vbComponents.Add(1); // 1 = vbext_ct_StdModule
                newModule.Name = moduleName;

                codeModule = newModule.CodeModule;
                codeModule.AddFromString(vbaCode);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                throw new InvalidOperationException(VbaTrustErrorMessage, comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref newModule);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Update(IExcelBatch batch, string moduleName, string vbaCode)
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
            dynamic? codeModule = null;
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
                    throw new InvalidOperationException($"Module '{moduleName}' not found. Use script-import to create it.");
                }

                codeModule = targetComponent.CodeModule;
                int lineCount = codeModule.CountOfLines;

                if (lineCount > 0)
                {
                    codeModule.DeleteLines(1, lineCount);
                }

                codeModule.AddFromString(vbaCode);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                throw new InvalidOperationException(VbaTrustErrorMessage, comEx);
            }
            finally
            {
                ComUtilities.Release(ref codeModule);
                ComUtilities.Release(ref targetComponent);
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }
}



