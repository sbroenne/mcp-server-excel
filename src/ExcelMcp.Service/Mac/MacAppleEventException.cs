namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacAppleEventException(int status, string operation, bool isExcelReply)
    : InvalidOperationException($"Excel Apple Event '{operation}' failed with OSStatus {status}.")
{
    internal int Status { get; } = status;
    internal string Operation { get; } = operation;
    internal bool IsExcelReply { get; } = isExcelReply;
}
