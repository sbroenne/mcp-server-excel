using System.Runtime.InteropServices;
using Microsoft.Win32;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands - Core data layer (no console output)
/// </summary>
public class ScriptCommands : IScriptCommands
{
    /// <summary>
    /// Check if VBA trust is enabled by reading registry
    /// </summary>
    private static bool IsVbaTrustEnabled()
    {
        try
        {
            // Try different Office versions
            string[] registryPaths = {
                @"Software\Microsoft\Office\16.0\Excel\Security",  // Office 2019/2021/365
                @"Software\Microsoft\Office\15.0\Excel\Security",  // Office 2013
                @"Software\Microsoft\Office\14.0\Excel\Security"   // Office 2010
            };

            foreach (string path in registryPaths)
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(path);
                    var value = key?.GetValue("AccessVBOM");
                    if (value != null && (int)value == 1)
                    {
                        return true;
                    }
                }
                catch { /* Try next path */ }
            }

            return false; // Assume not enabled if cannot read registry
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Creates VBA trust guidance result
    /// </summary>
    private static VbaTrustRequiredResult CreateVbaTrustGuidance()
    {
        return new VbaTrustRequiredResult
        {
            Success = false,
            ErrorMessage = "VBA trust access is not enabled",
            IsTrustEnabled = false,
            Explanation = "VBA operations require 'Trust access to the VBA project object model' to be enabled in Excel settings. This is a one-time setup that allows programmatic access to VBA code."
        };
    }

    /// <summary>
    /// Validate that file is macro-enabled (.xlsm) for VBA operations
    /// </summary>
    private static (bool IsValid, string? ErrorMessage) ValidateVbaFile(string filePath)
    {
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsm")
        {
            return (false, $"VBA operations require macro-enabled workbooks (.xlsm). Current file has extension: {extension}");
        }
        return (true, null);
    }

    /// <inheritdoc />
    public async Task<ScriptListResult> ListAsync(IExcelBatch batch)
    {
        var result = new ScriptListResult { FilePath = batch.WorkbookPath };

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
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
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

                        string typeStr = type switch
                        {
                            1 => "Module",
                            2 => "Class",
                            3 => "Form",
                            100 => "Document",
                            _ => $"Type{type}"
                        };

                        var procedures = new List<string>();
                        int moduleLineCount = 0;
                        try
                        {
                            codeModule = component.CodeModule;
                            moduleLineCount = codeModule.CountOfLines;

                            // Parse procedures from code
                            for (int line = 1; line <= moduleLineCount; line++)
                            {
                                string codeLine = codeModule.Lines[line, 1];
                                if (codeLine.TrimStart().StartsWith("Sub ") ||
                                    codeLine.TrimStart().StartsWith("Function ") ||
                                    codeLine.TrimStart().StartsWith("Public Sub ") ||
                                    codeLine.TrimStart().StartsWith("Public Function ") ||
                                    codeLine.TrimStart().StartsWith("Private Sub ") ||
                                    codeLine.TrimStart().StartsWith("Private Function "))
                                {
                                    string procName = ExtractProcedureName(codeLine);
                                    if (!string.IsNullOrEmpty(procName))
                                    {
                                        procedures.Add(procName);
                                    }
                                }
                            }
                        }
                        catch { }

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
                result.Success = false;
                result.ErrorMessage = "VBA trust access is not enabled";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing scripts: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref vbComponents);
                ComUtilities.Release(ref vbaProject);
            }
        });
    }

    /// <inheritdoc />
    public async Task<ScriptViewResult> ViewAsync(IExcelBatch batch, string moduleName)
    {
        var result = new ScriptViewResult { FilePath = batch.WorkbookPath, ModuleName = moduleName };

        var (isValid, validationError) = ValidateVbaFile(batch.WorkbookPath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (string.IsNullOrWhiteSpace(moduleName))
        {
            result.Success = false;
            result.ErrorMessage = "Module name cannot be empty";
            return result;
        }

        // Check VBA trust BEFORE attempting operation
        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? component = null;
            dynamic? codeModule = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
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
                        result.ModuleType = type switch
                        {
                            1 => "Module",
                            2 => "Class",
                            3 => "Form",
                            100 => "Document",
                            _ => $"Type{type}"
                        };

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
                            if (codeLine.TrimStart().StartsWith("Sub ") ||
                                codeLine.TrimStart().StartsWith("Function ") ||
                                codeLine.TrimStart().StartsWith("Public Sub ") ||
                                codeLine.TrimStart().StartsWith("Public Function ") ||
                                codeLine.TrimStart().StartsWith("Private Sub ") ||
                                codeLine.TrimStart().StartsWith("Private Function "))
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
                    result.Success = false;
                    result.ErrorMessage = $"Module '{moduleName}' not found in workbook";
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Module has {result.LineCount} lines and {result.Procedures.Count} procedure(s)",
                    "Use 'script-update' to modify the code",
                    "Use 'script-run' to execute procedures",
                    "Use 'script-export' to save code to file"
                };
                result.WorkflowHint = "VBA code viewed. Next, update or run procedures.";

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                result.Success = false;
                result.ErrorMessage = "VBA trust access is not enabled";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing script: {ex.Message}";
                return result;
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

    private static string ExtractProcedureName(string codeLine)
    {
        var parts = codeLine.Trim().Split(new[] { ' ', '(' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < parts.Length; i++)
        {
            if (parts[i] == "Sub" || parts[i] == "Function")
            {
                if (i + 1 < parts.Length)
                {
                    return parts[i + 1];
                }
            }
        }
        return string.Empty;
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(IExcelBatch batch, string moduleName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "script-export"
        };

        // Validate and normalize the output file path to prevent path traversal attacks
        try
        {
            outputFile = PathValidator.ValidateOutputFile(outputFile, nameof(outputFile), allowOverwrite: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid output file path: {ex.Message}";
            return result;
        }

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

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? targetComponent = null;
            dynamic? codeModule = null;
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
                    result.ErrorMessage = $"Script module '{moduleName}' not found";
                    return result;
                }

                codeModule = targetComponent.CodeModule;
                int lineCount = codeModule.CountOfLines;

                if (lineCount == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Module '{moduleName}' is empty";
                    return result;
                }

                string code = codeModule.Lines[1, lineCount];
                await File.WriteAllTextAsync(outputFile, code, ct);

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result.Success = false;
                result.ErrorMessage = "VBA trust access is not enabled";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting script: {ex.Message}";
                return result;
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

    /// <inheritdoc />
    public async Task<OperationResult> ImportAsync(IExcelBatch batch, string moduleName, string vbaFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "script-import"
        };

        // Validate and normalize the VBA file path to prevent path traversal attacks
        try
        {
            vbaFile = PathValidator.ValidateExistingFile(vbaFile, nameof(vbaFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid VBA file path: {ex.Message}";
            return result;
        }

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

        string vbaCode = await File.ReadAllTextAsync(vbaFile);

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? newModule = null;
            dynamic? codeModule = null;
            try
            {
                vbaProject = ctx.Book.VBProject;
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
                            result.Success = false;
                            result.ErrorMessage = $"Module '{moduleName}' already exists. Use script-update to modify it.";
                            return result;
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
                result.ErrorMessage = $"Error importing script: {ex.Message}";
                return result;
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
    public async Task<OperationResult> UpdateAsync(IExcelBatch batch, string moduleName, string vbaFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "script-update"
        };

        // Validate and normalize the VBA file path to prevent path traversal attacks
        try
        {
            vbaFile = PathValidator.ValidateExistingFile(vbaFile, nameof(vbaFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid VBA file path: {ex.Message}";
            return result;
        }

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

        string vbaCode = await File.ReadAllTextAsync(vbaFile);

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? vbaProject = null;
            dynamic? vbComponents = null;
            dynamic? targetComponent = null;
            dynamic? codeModule = null;
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
                    result.ErrorMessage = $"Module '{moduleName}' not found. Use script-import to create it.";
                    return result;
                }

                codeModule = targetComponent.CodeModule;
                int lineCount = codeModule.CountOfLines;

                if (lineCount > 0)
                {
                    codeModule.DeleteLines(1, lineCount);
                }

                codeModule.AddFromString(vbaCode);

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
                result.ErrorMessage = $"Error updating script: {ex.Message}";
                return result;
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

    /// <inheritdoc />
    public async Task<OperationResult> RunAsync(IExcelBatch batch, string procedureName, params string[] parameters)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "script-run"
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

        return await batch.ExecuteAsync(async (ctx, ct) =>
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
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string moduleName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "script-delete"
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

        return await batch.ExecuteAsync(async (ctx, ct) =>
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
