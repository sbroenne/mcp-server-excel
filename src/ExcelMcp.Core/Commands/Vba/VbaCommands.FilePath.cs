using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script operations - FilePath-based API implementations
/// </summary>
public partial class VbaCommands
{
    /// <inheritdoc />
    public async Task<VbaListResult> ListAsync(string filePath)
    {
        var result = new VbaListResult { FilePath = filePath };

        var (isValid, validationError) = ValidateVbaFile(filePath);
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
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? vbaProject = null;
                dynamic? vbComponents = null;
                try
                {
                    vbaProject = handle.Workbook.VBProject;
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
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    // Trust was disabled during operation
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error listing scripts: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbaProject);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<VbaViewResult> ViewAsync(string filePath, string moduleName)
    {
        var result = new VbaViewResult { FilePath = filePath, ModuleName = moduleName };

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? vbaProject = null;
                dynamic? vbComponents = null;
                dynamic? component = null;
                dynamic? codeModule = null;
                try
                {
                    vbaProject = handle.Workbook.VBProject;
                    vbComponents = vbaProject.VBComponents;

                    try
                    {
                        component = vbComponents.Item(moduleName);
                    }
                    catch
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Module '{moduleName}' not found";
                        return;
                    }

                    codeModule = component.CodeModule;
                    int lineCount = codeModule.CountOfLines;

                    if (lineCount > 0)
                    {
                        result.Code = codeModule.Lines[1, lineCount];
                        result.LineCount = lineCount;
                    }
                    else
                    {
                        result.Code = string.Empty;
                        result.LineCount = 0;
                    }

                    result.Success = true;
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error viewing module: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref codeModule);
                    ComUtilities.Release(ref component);
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbaProject);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(string filePath, string moduleName, string outputFile)
    {
        var result = new OperationResult { FilePath = filePath, Action = "export-vba" };

        var (isValid, validationError) = ValidateVbaFile(filePath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? vbaProject = null;
                dynamic? vbComponents = null;
                dynamic? component = null;
                try
                {
                    vbaProject = handle.Workbook.VBProject;
                    vbComponents = vbaProject.VBComponents;

                    try
                    {
                        component = vbComponents.Item(moduleName);
                    }
                    catch
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Module '{moduleName}' not found";
                        return;
                    }

                    component.Export(outputFile);
                    result.Success = true;
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error exporting module: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref component);
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbaProject);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> ImportAsync(string filePath, string moduleName, string vbaFile)
    {
        var result = new OperationResult { FilePath = filePath, Action = "import-vba" };

        var (isValid, validationError) = ValidateVbaFile(filePath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (!File.Exists(vbaFile))
        {
            result.Success = false;
            result.ErrorMessage = $"VBA file not found: {vbaFile}";
            return result;
        }

        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? vbaProject = null;
                dynamic? vbComponents = null;
                dynamic? imported = null;
                try
                {
                    vbaProject = handle.Workbook.VBProject;
                    vbComponents = vbaProject.VBComponents;

                    imported = vbComponents.Import(vbaFile);

                    if (!string.IsNullOrWhiteSpace(moduleName))
                    {
                        imported.Name = moduleName;
                    }

                    result.Success = true;
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error importing module: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref imported);
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbaProject);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(string filePath, string moduleName, string vbaFile)
    {
        var result = new OperationResult { FilePath = filePath, Action = "update-vba" };

        var (isValid, validationError) = ValidateVbaFile(filePath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (!File.Exists(vbaFile))
        {
            result.Success = false;
            result.ErrorMessage = $"VBA file not found: {vbaFile}";
            return result;
        }

        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? vbaProject = null;
                dynamic? vbComponents = null;
                dynamic? component = null;
                dynamic? codeModule = null;
                try
                {
                    vbaProject = handle.Workbook.VBProject;
                    vbComponents = vbaProject.VBComponents;

                    try
                    {
                        component = vbComponents.Item(moduleName);
                    }
                    catch
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Module '{moduleName}' not found";
                        return;
                    }

                    string newCode = File.ReadAllText(vbaFile);
                    codeModule = component.CodeModule;
                    int lineCount = codeModule.CountOfLines;

                    if (lineCount > 0)
                    {
                        codeModule.DeleteLines(1, lineCount);
                    }

                    if (!string.IsNullOrWhiteSpace(newCode))
                    {
                        codeModule.AddFromString(newCode);
                    }

                    result.Success = true;
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error updating module: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref codeModule);
                    ComUtilities.Release(ref component);
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbaProject);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> RunAsync(string filePath, string procedureName, TimeSpan? timeout, params string[] parameters)
    {
        var result = new OperationResult { FilePath = filePath, Action = "run-vba" };

        var (isValid, validationError) = ValidateVbaFile(filePath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                try
                {
                    if (parameters != null && parameters.Length > 0)
                    {
                        object[] args = new object[parameters.Length];
                        for (int i = 0; i < parameters.Length; i++)
                        {
                            args[i] = parameters[i];
                        }
                        handle.Application.Run(procedureName, args[0], args.Length > 1 ? args[1] : System.Reflection.Missing.Value);
                    }
                    else
                    {
                        handle.Application.Run(procedureName);
                    }

                    result.Success = true;
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error running procedure: {ex.Message}";
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string moduleName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "delete-vba" };

        var (isValid, validationError) = ValidateVbaFile(filePath);
        if (!isValid)
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (!IsVbaTrustEnabled())
        {
            var trustGuidance = CreateVbaTrustGuidance();
            result.Success = false;
            result.ErrorMessage = trustGuidance.ErrorMessage;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? vbaProject = null;
                dynamic? vbComponents = null;
                dynamic? component = null;
                try
                {
                    vbaProject = handle.Workbook.VBProject;
                    vbComponents = vbaProject.VBComponents;

                    try
                    {
                        component = vbComponents.Item(moduleName);
                    }
                    catch
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Module '{moduleName}' not found";
                        return;
                    }

                    vbComponents.Remove(component);
                    result.Success = true;
                }
                catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                                 comEx.ErrorCode == unchecked((int)0x800A03EC))
                {
                    result.Success = false;
                    result.ErrorMessage = "VBA trust access is not enabled";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error deleting module: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref component);
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbaProject);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }
}
