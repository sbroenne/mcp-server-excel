namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class ExcelProcessPersistenceException : InvalidOperationException
{
    public ExcelProcessPersistenceException(ExcelProcessIdentity identity, Exception innerException)
        : base(
            $"Failed to durably persist ownership for Excel process {identity.ProcessId}.",
            innerException)
    {
        Identity = identity;
    }

    public ExcelProcessIdentity Identity { get; }
}
