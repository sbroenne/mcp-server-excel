using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    /// <inheritdoc />
    public DocumentInspectionResult InspectDocument(IExcelBatch batch) =>
        batch.Execute((ctx, ct) =>
        {
            Office.DocumentInspectors? inspectors = null;
            List<DocumentInspectorFinding> findings = [];
            bool savedBefore = ctx.Book.Saved;
            try
            {
                inspectors = ctx.Book.DocumentInspectors;
                int count = Convert.ToInt32(inspectors.Count);
                for (int index = 1; index <= count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Office.DocumentInspector? inspector = null;
                    try
                    {
                        inspector = inspectors[index];
                        string name = inspector.Name;
                        var outcome = InvokeDocumentInspector(inspector, fix: false);
                        findings.Add(new DocumentInspectorFinding(index, name, outcome.StatusCode, outcome.Results));
                    }
                    finally
                    {
                        ComUtilities.Release(ref inspector);
                    }
                }
                var result = CreateDocumentInspectionResult(findings);
                result.FilePath = batch.WorkbookPath;
                result.SavedBefore = savedBefore;
                result.SavedAfter = ctx.Book.Saved;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref inspectors);
            }
        });

    /// <inheritdoc />
    public DocumentInspectorFixResult FixDocumentInspector(
        IExcelBatch batch, int inspectorIndex, string inspectorName)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(inspectorIndex, 1);
        ArgumentException.ThrowIfNullOrWhiteSpace(inspectorName);
        return batch.Execute((ctx, ct) =>
        {
            ct.ThrowIfCancellationRequested();
            if (ctx.Book.ReadOnly)
                throw new InvalidOperationException(
                    "Cannot fix a document inspector: Excel opened this workbook read-only. The workbook has not been changed.");
            if (ctx.Book.ProtectStructure || ctx.Book.ProtectWindows)
                throw new InvalidOperationException(
                    "Cannot fix a document inspector on a structure/window-protected workbook. The workbook has not been changed.");
            Office.DocumentInspectors? inspectors = null;
            Office.DocumentInspector? inspector = null;
            try
            {
                inspectors = ctx.Book.DocumentInspectors;
                int count = Convert.ToInt32(inspectors.Count);
                if (inspectorIndex > count)
                    throw new ArgumentOutOfRangeException(nameof(inspectorIndex),
                        $"inspectorIndex must select an available module, 1 through {count}. The workbook has not been changed.");
                inspector = inspectors[inspectorIndex];
                string actualName = inspector.Name;
                if (!string.Equals(actualName, inspectorName, StringComparison.Ordinal))
                    throw new ArgumentException(
                        $"Inspector name does not match index {inspectorIndex}. Run inspect-document again and use its exact name and index. " +
                        "The workbook has not been changed.", nameof(inspectorName));

                ct.ThrowIfCancellationRequested();
                DocumentInspectorOutcome fix;
                DocumentInspectorOutcome inspection;
                try
                {
                    fix = InvokeDocumentInspector(inspector, fix: true);
                    ct.ThrowIfCancellationRequested();
                    inspection = InvokeDocumentInspector(inspector, fix: false);
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException(
                        $"Document Inspector '{actualName}' (index {inspectorIndex}) failed during fix or reinspection: {ex.Message} " +
                        "The workbook may already have changed. No rollback or automatic save was performed; inspect its current state before retrying.",
                        ex);
                }
                var result = CreateDocumentInspectorFixResult(inspectorIndex, actualName, fix, inspection);
                result.FilePath = batch.WorkbookPath;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref inspector);
                ComUtilities.Release(ref inspectors);
            }
        });
    }

    private static DocumentInspectorOutcome InvokeDocumentInspector(Office.DocumentInspector inspector, bool fix)
    {
        Office.MsoDocInspectorStatus status;
        string results;
        if (fix)
            inspector.Fix(out status, out results);
        else
            inspector.Inspect(out status, out results);
        return new DocumentInspectorOutcome((int)status, results);
    }

    internal static DocumentInspectionResult CreateDocumentInspectionResult(List<DocumentInspectorFinding> findings)
    {
        var failures = findings.Where(item => item.StatusCode is not (0 or 1)).ToArray();
        return new DocumentInspectionResult
        {
            Inspectors = findings,
            Success = findings.Count > 0 && failures.Length == 0,
            ErrorMessage = findings.Count == 0
                ? "Excel exposed no Document Inspector modules. No document inspection was performed."
                : failures.Length == 0 ? null
                : "Document inspection failed for " + string.Join("; ", failures.Select(item =>
                    $"'{item.Name}' (index {item.Index}, status {item.StatusCode}): {item.Results}"))
        };
    }

    internal static DocumentInspectorFixResult CreateDocumentInspectorFixResult(
        int inspectorIndex, string inspectorName, DocumentInspectorOutcome fix, DocumentInspectorOutcome inspection)
    {
        bool success = fix.StatusCode == 0 && inspection.StatusCode == 0;
        return new DocumentInspectorFixResult
        {
            Action = "fix-document-inspector",
            InspectorIndex = inspectorIndex,
            InspectorName = inspectorName,
            Fix = fix,
            Inspection = inspection,
            Success = success,
            Message = success ? "The selected inspector's fix completed and reinspection found no remaining issues. Saving remains explicit." : null,
            ErrorMessage = success ? null
                : $"Document Inspector '{inspectorName}' (index {inspectorIndex}) did not confirm a complete fix. " +
                $"Fix status {fix.StatusCode}: {fix.Results}; reinspection status {inspection.StatusCode}: {inspection.Results}. " +
                "The workbook may already have changed. No rollback or automatic save was performed; inspect its current state before retrying."
        };
    }
}
