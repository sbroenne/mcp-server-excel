using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands - FilePath-based API
/// </summary>
public partial class VbaCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
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
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error opening workbook: {ex.Message}";
        }

        return result;
    }

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    public async Task<VbaViewResult> ViewAsync(string filePath, string moduleName)
    {
        // Delegate to batch-based implementation for now
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ViewAsync(batch, moduleName);
    }

    /// <summary>
    /// Exports VBA module code to a file
    /// </summary>
    public async Task<OperationResult> ExportAsync(string filePath, string moduleName, string outputFile)
    {
        // Delegate to batch-based implementation for now
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ExportAsync(batch, moduleName, outputFile);
    }

    /// <summary>
    /// Imports VBA code from a file to create a new module
    /// </summary>
    public async Task<OperationResult> ImportAsync(string filePath, string moduleName, string vbaFile)
    {
        // Delegate to batch-based implementation for now
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await ImportAsync(batch, moduleName, vbaFile);
    }

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    public async Task<OperationResult> UpdateAsync(string filePath, string moduleName, string vbaFile)
    {
        // Delegate to batch-based implementation for now
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await UpdateAsync(batch, moduleName, vbaFile);
    }

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    public async Task<OperationResult> RunAsync(string filePath, string procedureName, TimeSpan? timeout, params string[] parameters)
    {
        // Delegate to batch-based implementation for now
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await RunAsync(batch, procedureName, timeout, parameters);
    }

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    public async Task<OperationResult> DeleteAsync(string filePath, string moduleName)
    {
        // Delegate to batch-based implementation for now
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        return await DeleteAsync(batch, moduleName);
    }
}
