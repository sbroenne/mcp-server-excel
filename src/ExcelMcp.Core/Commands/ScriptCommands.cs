using System.Runtime.InteropServices;
using Microsoft.Win32;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

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
    /// Check if VBA project access is trusted and available
    /// </summary>
    private static (bool IsTrusted, string? ErrorMessage) CheckVbaAccessTrust(string filePath)
    {
        try
        {
            bool isTrusted = false;
            string? errorMessage = null;

            WithExcel(filePath, false, (excel, workbook) =>
            {
                try
                {
                    dynamic vbProject = workbook.VBProject;
                    int componentCount = vbProject.VBComponents.Count;
                    isTrusted = true;
                    return 0;
                }
                catch (COMException comEx)
                {
                    if (comEx.ErrorCode == unchecked((int)0x800A03EC))
                    {
                        errorMessage = "Programmatic access to VBA project is not trusted. Run setup-vba-trust command.";
                    }
                    else
                    {
                        errorMessage = $"VBA COM Error: 0x{comEx.ErrorCode:X8} - {comEx.Message}";
                    }
                    return 1;
                }
                catch (Exception ex)
                {
                    errorMessage = $"VBA Access Error: {ex.Message}";
                    return 1;
                }
            });

            return (isTrusted, errorMessage);
        }
        catch (Exception ex)
        {
            return (false, $"Error checking VBA access: {ex.Message}");
        }
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
    public ScriptListResult List(string filePath)
    {
        var result = new ScriptListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic vbaProject = workbook.VBProject;
                dynamic vbComponents = vbaProject.VBComponents;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic component = vbComponents.Item(i);
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
                        dynamic codeModule = component.CodeModule;
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

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result.Success = false;
                result.ErrorMessage = "VBA trust access is not enabled";
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing scripts: {ex.Message}";
                return 1;
            }
        });

        return result;
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
    public async Task<OperationResult> Export(string filePath, string moduleName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "script-export"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

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

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic vbaProject = workbook.VBProject;
                dynamic vbComponents = vbaProject.VBComponents;
                dynamic? targetComponent = null;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic component = vbComponents.Item(i);
                    if (component.Name == moduleName)
                    {
                        targetComponent = component;
                        break;
                    }
                }

                if (targetComponent == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Script module '{moduleName}' not found";
                    return 1;
                }

                dynamic codeModule = targetComponent.CodeModule;
                int lineCount = codeModule.CountOfLines;

                if (lineCount == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Module '{moduleName}' is empty";
                    return 1;
                }

                string code = codeModule.Lines[1, lineCount];
                File.WriteAllText(outputFile, code);

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = filePath;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting script: {ex.Message}";
                return 1;
            }
        });

        return await Task.FromResult(result);
    }

    /// <inheritdoc />
    public async Task<OperationResult> Import(string filePath, string moduleName, string vbaFile)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "script-import"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

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

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic vbaProject = workbook.VBProject;
                dynamic vbComponents = vbaProject.VBComponents;

                // Check if module already exists
                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic component = vbComponents.Item(i);
                    if (component.Name == moduleName)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Module '{moduleName}' already exists. Use script-update to modify it.";
                        return 1;
                    }
                }

                // Add new module
                dynamic newModule = vbComponents.Add(1); // 1 = vbext_ct_StdModule
                newModule.Name = moduleName;

                dynamic codeModule = newModule.CodeModule;
                codeModule.AddFromString(vbaCode);

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = filePath;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error importing script: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> Update(string filePath, string moduleName, string vbaFile)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "script-update"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

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

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic vbaProject = workbook.VBProject;
                dynamic vbComponents = vbaProject.VBComponents;
                dynamic? targetComponent = null;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic component = vbComponents.Item(i);
                    if (component.Name == moduleName)
                    {
                        targetComponent = component;
                        break;
                    }
                }

                if (targetComponent == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Module '{moduleName}' not found. Use script-import to create it.";
                    return 1;
                }

                dynamic codeModule = targetComponent.CodeModule;
                int lineCount = codeModule.CountOfLines;

                if (lineCount > 0)
                {
                    codeModule.DeleteLines(1, lineCount);
                }

                codeModule.AddFromString(vbaCode);

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = filePath;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating script: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Run(string filePath, string procedureName, params string[] parameters)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "script-run"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                if (parameters.Length == 0)
                {
                    excel.Run(procedureName);
                }
                else
                {
                    object[] paramObjects = parameters.Cast<object>().ToArray();
                    excel.Run(procedureName, paramObjects);
                }

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = filePath;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error running procedure '{procedureName}': {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Delete(string filePath, string moduleName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "script-delete"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        var (isValid, validationError) = ValidateVbaFile(filePath);
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

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic vbaProject = workbook.VBProject;
                dynamic vbComponents = vbaProject.VBComponents;
                dynamic? targetComponent = null;

                for (int i = 1; i <= vbComponents.Count; i++)
                {
                    dynamic component = vbComponents.Item(i);
                    if (component.Name == moduleName)
                    {
                        targetComponent = component;
                        break;
                    }
                }

                if (targetComponent == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Module '{moduleName}' not found";
                    return 1;
                }

                vbComponents.Remove(targetComponent);

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase) ||
                                             comEx.ErrorCode == unchecked((int)0x800A03EC))
            {
                // Trust was disabled during operation
                result = CreateVbaTrustGuidance();
                result.FilePath = filePath;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting module: {ex.Message}";
                return 1;
            }
        });

        return result;
    }
}
