namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class OperationCleanup
{
    internal static void Run(Action cleanup, Exception? operationFailure)
    {
        try
        {
            cleanup();
        }
        catch (Exception cleanupFailure) when (operationFailure is not null)
        {
            var failure = new AggregateException("The operation and restoring its temporary state both failed.",
                operationFailure, cleanupFailure);
            if (operationFailure is OperationCanceledException canceled)
                throw new OperationCanceledException(canceled.Message, failure, canceled.CancellationToken);
            throw failure;
        }
    }
}
