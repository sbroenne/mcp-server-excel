using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Reports failures that happen after Excel already created a workbook object.
/// The object is intentionally kept so the caller can inspect it and decide what to do.
/// </summary>
internal static class CreatedObjectFailure
{
    /// <summary>
    /// Returns false for cancellation and lost-Excel failures, which must keep their original shape.
    /// </summary>
    internal static bool CanReport(Exception exception)
    {
        if (exception is OperationCanceledException)
        {
            return false;
        }

        for (Exception? current = exception; current != null; current = current.InnerException)
        {
            if (current is COMException com
                && com.HResult is ResiliencePipelines.RPC_E_CALL_FAILED
                    or ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE
                    or ResiliencePipelines.RPC_E_DISCONNECTED)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Builds an error that names the object left in the workbook and where to find it.
    /// </summary>
    internal static InvalidOperationException Create(
        string objectKind,
        string objectName,
        string sheetName,
        string failedStep,
        Exception cause) =>
        CreateDescribed(
            failedStep,
            $"{objectKind} '{objectName}' on sheet '{sheetName}'",
            $"The {objectKind} '{objectName}' remains on sheet '{sheetName}'; inspect it and remove it if appropriate.",
            cause);

    /// <summary>
    /// Builds an error for a created object whose location is not a single sheet.
    /// </summary>
    internal static InvalidOperationException CreateDescribed(
        string failedStep,
        string createdObject,
        string remainingState,
        Exception cause)
    {
        var message = $"{failedStep} failed after Excel created {createdObject}: {cause.Message} {remainingState}";
        return cause is OperationFailureException categorized
            ? new OperationFailureException(categorized.ErrorCategory, message, cause)
            : new InvalidOperationException(message, cause);
    }
}
