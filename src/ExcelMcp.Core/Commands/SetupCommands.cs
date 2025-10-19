using Microsoft.Win32;
using Sbroenne.ExcelMcp.Core.Models;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Setup and configuration commands for ExcelCLI
/// </summary>
public class SetupCommands : ISetupCommands
{
    /// <inheritdoc />
    public VbaTrustResult EnableVbaTrust()
    {
        try
        {
            // Try different Office versions and architectures
            string[] registryPaths = {
                @"SOFTWARE\Microsoft\Office\16.0\Excel\Security",  // Office 2019/2021/365
                @"SOFTWARE\Microsoft\Office\15.0\Excel\Security",  // Office 2013
                @"SOFTWARE\Microsoft\Office\14.0\Excel\Security",  // Office 2010
                @"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Excel\Security",  // 32-bit on 64-bit
                @"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Excel\Security",
                @"SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Excel\Security"
            };

            var result = new VbaTrustResult();
            
            foreach (string path in registryPaths)
            {
                try
                {
                    using (RegistryKey? key = Registry.CurrentUser.CreateSubKey(path))
                    {
                        if (key != null)
                        {
                            // Set AccessVBOM = 1 to trust VBA project access
                            key.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
                            result.RegistryPathsSet.Add(path);
                        }
                    }
                }
                catch
                {
                    // Skip paths that don't exist or can't be accessed
                }
            }

            if (result.RegistryPathsSet.Count > 0)
            {
                result.Success = true;
                result.IsTrusted = true;
                result.ManualInstructions = "You may need to restart Excel for changes to take effect.";
            }
            else
            {
                result.Success = false;
                result.IsTrusted = false;
                result.ErrorMessage = "Could not find Excel registry keys to modify.";
                result.ManualInstructions = "File → Options → Trust Center → Trust Center Settings → Macro Settings\nCheck 'Trust access to the VBA project object model'";
            }
            
            return result;
        }
        catch (Exception ex)
        {
            return new VbaTrustResult
            {
                Success = false,
                IsTrusted = false,
                ErrorMessage = ex.Message,
                ManualInstructions = "File → Options → Trust Center → Trust Center Settings → Macro Settings\nCheck 'Trust access to the VBA project object model'"
            };
        }
    }

    /// <inheritdoc />
    public VbaTrustResult CheckVbaTrust(string testFilePath)
    {
        if (string.IsNullOrEmpty(testFilePath))
        {
            return new VbaTrustResult
            {
                Success = false,
                IsTrusted = false,
                ErrorMessage = "Test file path is required",
                FilePath = testFilePath
            };
        }

        if (!File.Exists(testFilePath))
        {
            return new VbaTrustResult
            {
                Success = false,
                IsTrusted = false,
                ErrorMessage = $"Test file not found: {testFilePath}",
                FilePath = testFilePath
            };
        }

        try
        {
            var result = new VbaTrustResult { FilePath = testFilePath };
            
            int exitCode = WithExcel(testFilePath, false, (excel, workbook) =>
            {
                try
                {
                    dynamic vbProject = workbook.VBProject;
                    result.ComponentCount = vbProject.VBComponents.Count;
                    result.IsTrusted = true;
                    result.Success = true;
                    return 0;
                }
                catch (Exception ex)
                {
                    result.IsTrusted = false;
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                    result.ManualInstructions = "Run 'setup-vba-trust' or manually: File → Options → Trust Center → Trust Center Settings → Macro Settings\nCheck 'Trust access to the VBA project object model'";
                    return 1;
                }
            });
            
            return result;
        }
        catch (Exception ex)
        {
            return new VbaTrustResult
            {
                Success = false,
                IsTrusted = false,
                ErrorMessage = $"Error testing VBA access: {ex.Message}",
                FilePath = testFilePath
            };
        }
    }
}
