namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class ExcelApplicationLease : IDisposable
{
    private readonly Mutex _mutex;
    private bool _disposed;

    internal ExcelApplicationLease(ExcelProcessIdentity identity)
    {
        _mutex = new Mutex(false,
            $@"Local\ExcelMcp.Application.{identity.ProcessId}.{identity.StartedAtUtcFileTime}");
        try
        {
            bool acquired;
            try
            {
                acquired = _mutex.WaitOne(TimeSpan.Zero);
            }
            catch (AbandonedMutexException)
            {
                acquired = true;
            }

            if (!acquired)
                throw new InvalidOperationException(
                    "This Excel application is already managed by another active ExcelMcp host.");
        }
        catch
        {
            _mutex.Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        try
        {
            _mutex.ReleaseMutex();
        }
        finally
        {
            _disposed = true;
            _mutex.Dispose();
        }
    }
}
