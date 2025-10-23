using System.Runtime.InteropServices;
using Microsoft.Extensions.ObjectPool;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Pooling policy for Excel COM instances. Defines how to create, reset, and destroy pooled Excel instances.
/// Used by Microsoft.Extensions.ObjectPool for battle-tested object lifecycle management.
/// </summary>
internal class ExcelInstancePoolPolicy : IPooledObjectPolicy<PooledExcelInstance>
{
    private readonly TimeSpan _idleTimeout;

    public ExcelInstancePoolPolicy(TimeSpan idleTimeout)
    {
        _idleTimeout = idleTimeout;
    }

    /// <summary>
    /// Creates a new Excel COM instance for the pool.
    /// </summary>
    public PooledExcelInstance Create()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed on this system");

        dynamic excel = Activator.CreateInstance(excelType)
            ?? throw new InvalidOperationException("Failed to create Excel instance");

        // Configure Excel for automation
        excel.Visible = false;
        excel.DisplayAlerts = false;
        excel.EnableEvents = false;

        return new PooledExcelInstance
        {
            Excel = excel,
            Workbook = null,
            LastUsed = DateTime.UtcNow,
            Lock = new object()
        };
    }

    /// <summary>
    /// Prepares instance for return to pool. Closes workbook but keeps Excel alive.
    /// Returns false if instance should be destroyed instead of returned to pool.
    /// </summary>
    public bool Return(PooledExcelInstance instance)
    {
        try
        {
            // Check if instance has been idle too long
            if (DateTime.UtcNow - instance.LastUsed > _idleTimeout)
            {
                return false; // Destroy idle instance
            }

            // Close workbook if open (but keep Excel instance alive)
            if (instance.Workbook != null)
            {
                try
                {
                    instance.Workbook.Close(false);
                    Marshal.ReleaseComObject(instance.Workbook);
                }
                catch
                {
                    // Ignore cleanup errors
                }
                finally
                {
                    instance.Workbook = null;
                }
            }

            // Update last used timestamp
            instance.LastUsed = DateTime.UtcNow;
            return true; // Instance is healthy, return to pool
        }
        catch
        {
            return false; // Destroy unhealthy instance
        }
    }
}
