using Microsoft.Win32;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// TEST INFRASTRUCTURE ONLY - Temporarily modifies VBA trust registry setting 
/// for isolated test execution. NEVER expose this to end users!
/// 
/// This class is INTERNAL and located in the test project only.
/// It should NEVER be referenced by Core/CLI/MCP production code.
/// </summary>
internal sealed class TestVbaTrustScope : IDisposable
{
    private readonly bool _wasEnabled;
    private bool _isDisposed;
    
    public TestVbaTrustScope()
    {
        _wasEnabled = IsVbaTrustEnabled();
        if (!_wasEnabled)
        {
            EnableVbaTrust();
            Thread.Sleep(150); // Registry propagation delay
            
            if (!IsVbaTrustEnabled())
                throw new InvalidOperationException("Test setup failed: Could not enable VBA trust");
        }
    }
    
    public void Dispose()
    {
        if (!_isDisposed && !_wasEnabled)
        {
            try { DisableVbaTrust(); }
            catch (Exception ex)
            {
                // Log but don't throw in Dispose
                Console.Error.WriteLine($"Test cleanup warning: Could not disable VBA trust: {ex.Message}");
            }
            finally 
            { 
                _isDisposed = true;
                GC.SuppressFinalize(this);
            }
        }
    }
    
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
            
            return false;
        }
        catch
        {
            return false;
        }
    }
    
    private static void EnableVbaTrust()
    {
        try
        {
            // Try different Office versions
            string[] registryPaths = {
                @"Software\Microsoft\Office\16.0\Excel\Security",
                @"Software\Microsoft\Office\15.0\Excel\Security",
                @"Software\Microsoft\Office\14.0\Excel\Security"
            };
            
            foreach (string path in registryPaths)
            {
                try
                {
                    using var key = Registry.CurrentUser.CreateSubKey(path);
                    key?.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
                }
                catch { /* Try next path */ }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to enable VBA trust for test: {ex.Message}", ex);
        }
    }
    
    private static void DisableVbaTrust()
    {
        try
        {
            // Try different Office versions
            string[] registryPaths = {
                @"Software\Microsoft\Office\16.0\Excel\Security",
                @"Software\Microsoft\Office\15.0\Excel\Security",
                @"Software\Microsoft\Office\14.0\Excel\Security"
            };
            
            foreach (string path in registryPaths)
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(path, writable: true);
                    key?.SetValue("AccessVBOM", 0, RegistryValueKind.DWord);
                }
                catch { /* Try next path */ }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to disable VBA trust after test: {ex.Message}", ex);
        }
    }
}
