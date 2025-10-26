using System.Runtime.InteropServices;
using System.Threading.Channels;
using Sbroenne.ExcelMcp.Core.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Executes Excel COM operations on a dedicated STA thread with OLE message filter.
/// Ensures proper COM apartment state for all Excel automation.
/// </summary>
internal static class ExcelStaExecutor
{
    /// <summary>
    /// Executes an Excel operation on a dedicated STA thread.
    /// Ensures proper COM apartment state and OLE message filter registration.
    /// </summary>
    public static async Task<T> ExecuteOnStaThreadAsync<T>(
        Func<Task<T>> operation,
        CancellationToken cancellationToken = default)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                // CRITICAL: Register OLE message filter for Excel busy handling
                OleMessageFilter.Register();

                // Execute operation on STA thread
                var result = operation().GetAwaiter().GetResult();
                tcs.SetResult(result);
            }
            catch (OperationCanceledException oce)
            {
                tcs.TrySetCanceled(oce.CancellationToken);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        })
        {
            IsBackground = true,
            Name = "ExcelSTA"
        };

        // CRITICAL: Set STA apartment state before starting thread
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        // Support cancellation
        using var registration = cancellationToken.Register(() => tcs.TrySetCanceled(cancellationToken));

        return await tcs.Task;
    }
}
