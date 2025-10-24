using Microsoft.Win32;
using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

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
    public VbaTrustResult DisableVbaTrust()
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
                    using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(path, writable: true))
                    {
                        if (key != null)
                        {
                            // Set AccessVBOM = 0 to disable VBA project access
                            key.SetValue("AccessVBOM", 0, RegistryValueKind.DWord);
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
                result.IsTrusted = false;
                result.ManualInstructions = "VBA trust has been disabled. Restart Excel for changes to take effect.";
            }
            else
            {
                result.Success = false;
                result.IsTrusted = false;
                result.ErrorMessage = "Could not find Excel registry keys to modify.";
            }

            return result;
        }
        catch (Exception ex)
        {
            return new VbaTrustResult
            {
                Success = false,
                IsTrusted = false,
                ErrorMessage = ex.Message
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

            int exitCode = ExcelSession.Execute(testFilePath, false, (excel, workbook) =>
            {
                dynamic? vbProject = null;
                dynamic? vbComponents = null;
                try
                {
                    vbProject = workbook.VBProject;
                    vbComponents = vbProject.VBComponents;
                    result.ComponentCount = vbComponents.Count;
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
                finally
                {
                    ComUtilities.Release(ref vbComponents);
                    ComUtilities.Release(ref vbProject);
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
