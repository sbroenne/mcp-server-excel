using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class PowerQueryCommands
{
    internal static bool CanReportMutationFailure(Exception exception)
    {
        if (!CreatedObjectFailure.CanReport(exception))
            return false;

        for (Exception? current = exception; current != null; current = current.InnerException)
        {
            if (current is TimeoutException or OperationCanceledException or InvalidComObjectException
                or OperationFailureException { ErrorCategory: OperationFailureCategory.Cancelled })
                return false;
        }

        return true;
    }

    internal static PowerQueryCommandException CreateMutationFailure(
        Exception cause,
        string queryName,
        bool queryCreated,
        Action<PowerQueryPartialState> inspectState)
    {
        var state = new PowerQueryPartialState
        {
            QueryName = queryName,
            QueryCreated = queryCreated ? true : null,
            MCodeUpdated = queryCreated ? null : true
        };
        try
        {
            inspectState(state);
        }
        catch (Exception inspectionError) when (CanReportMutationFailure(inspectionError))
        {
            state.StateInspectionError = $"{inspectionError.GetType().Name}: {inspectionError.Message}";
        }

        var mutation = queryCreated
            ? $"Query '{queryName}' was created, but loading failed"
            : $"M code for query '{queryName}' was updated, but refreshing failed";
        var remaining = state.QueryExists switch
        {
            true when state.LoadMode is not null =>
                $"The query remains with load mode '{state.LoadMode}'" +
                (state.TargetSheet is null ? "." : $" on worksheet '{state.TargetSheet}'."),
            true => "The query remains; its load configuration could not be inspected.",
            false => "The query was not found during post-failure inspection.",
            null => "The surviving query state could not be inspected."
        };
        var inspection = state.StateInspectionError is null
            ? string.Empty
            : $" State inspection failed: {state.StateInspectionError}";
        var message = $"{mutation}: {cause.Message} {remaining}{inspection} " +
            "Changes were not rolled back. Inspect with list, view, or get-load-config; " +
            "use update to correct the stored M code, refresh for an existing load, " +
            "load-to to configure a destination, or delete only if removal is intended.";
        var category = OperationFailureClassifier.Classify(cause);
        if (cause is COMException && category == "ComInterop"
            && TryWrapPowerQueryException(cause, out var classified))
        {
            category = classified!.ErrorCategory;
        }
        return new PowerQueryCommandException(message, category, cause, state);
    }

    private static void InspectMutationState(
        Excel.Workbook workbook,
        PowerQueryPartialState state,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Excel.WorkbookQuery? query = null;
        try
        {
            query = PowerQuery.PowerQueryHelpers.FindQueryByExactName(workbook, state.QueryName);
            state.QueryExists = query != null;
            if (query == null)
                return;

            var loadState = DetectLoadState(workbook, state.QueryName, cancellationToken);
            state.LoadMode = loadState.LoadMode;
            state.TargetSheet = loadState.TargetSheet;
            state.IsLoadedToDataModel = loadState.IsLoadedToDataModel;
        }
        finally
        {
            ComUtilities.Release(ref query);
        }
    }
}
