using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Sbroenne.ExcelMcp.Service.Safety;

namespace Sbroenne.ExcelMcp.Service.Workflow;

/// <summary>
/// Owns the deliberately narrow compatibility contract for one-STA workflow execution.
/// A plan is selected atomically before dispatch; no fast prefix is ever retried through
/// the sequential executor.
/// </summary>
internal static class WorkflowFastPathPolicy
{
    internal const int MaximumFastOperations = 64;

    internal static readonly string[] CompatibleCategories =
        ["range", "rangeedit", "rangeformat", "rangelink", "sheet", "sheetstyle", "reportformat"];

    private static readonly HashSet<string> CompatibleCommands = new(StringComparer.Ordinal)
    {
        "range.get-values", "range.set-values", "range.get-formulas", "range.set-formulas",
        "range.validate-formulas", "range.clear-all", "range.clear-contents", "range.clear-formats",
        "range.get-number-formats", "range.set-number-format", "range.set-number-formats",
        "range.get-used-range", "range.get-current-region", "range.get-info",
        "rangeedit.insert-cells", "rangeedit.delete-cells", "rangeedit.insert-rows",
        "rangeedit.delete-rows", "rangeedit.insert-columns", "rangeedit.delete-columns",
        "rangeedit.find", "rangeedit.replace", "rangeedit.sort",
        "rangeformat.set-style", "rangeformat.get-style", "rangeformat.format-range",
        "rangeformat.format-ranges", "rangeformat.validate-range", "rangeformat.get-validation",
        "rangeformat.remove-validation", "rangeformat.auto-fit-columns", "rangeformat.auto-fit-rows",
        "rangeformat.merge-cells", "rangeformat.unmerge-cells", "rangeformat.get-merge-info",
        "rangeformat.set-column-width", "rangeformat.set-row-height",
        "rangelink.add-hyperlink", "rangelink.remove-hyperlink", "rangelink.list-hyperlinks",
        "rangelink.get-hyperlink", "rangelink.set-cell-lock", "rangelink.get-cell-lock",
        "sheet.list", "sheet.create", "sheet.rename", "sheet.copy", "sheet.delete", "sheet.move",
        "sheetstyle.set-tab-color", "sheetstyle.get-tab-color", "sheetstyle.clear-tab-color",
        "sheetstyle.set-visibility", "sheetstyle.get-visibility", "sheetstyle.show",
        "sheetstyle.hide", "sheetstyle.very-hide",
        "reportformat.apply", "reportformat.get-state",
    };

    internal static string? GetFallbackReason(
        WorkflowPlanRequest plan,
        SessionSafetyConfiguration safetyConfiguration,
        bool sharedCheckpointRequested)
    {
        if (!plan.FastMode)
        {
            return "fast_mode_disabled";
        }

        if (plan.Operations.Count > MaximumFastOperations)
        {
            return $"operation_count_exceeds_fast_limit:{MaximumFastOperations}";
        }

        if (sharedCheckpointRequested)
        {
            return "checkpoint_required";
        }

        if (safetyConfiguration.UsesSafetyWorkflow)
        {
            return "safety_workflow_enabled";
        }

        foreach (var operation in plan.Operations)
        {
            if (!CompatibleCommands.Contains(operation.Command))
            {
                return $"incompatible_command:{operation.Command}";
            }

            if (UsesExternalValueFile(operation))
            {
                return $"external_file_input:{operation.Command}";
            }
        }

        return null;
    }

    private static bool UsesExternalValueFile(ServiceBatchOperation operation)
    {
        if (operation.Command is not ("range.set-values" or "range.set-formulas") ||
            operation.Args is not { ValueKind: JsonValueKind.Object } args)
        {
            return false;
        }

        string propertyName = operation.Command == "range.set-values" ? "valuesFile" : "formulasFile";
        return args.TryGetProperty(propertyName, out var file) &&
            file.ValueKind == JsonValueKind.String &&
            !string.IsNullOrWhiteSpace(file.GetString());
    }
}
