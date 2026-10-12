using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

internal static class PivotTableSnapshotHelpers
{
    internal const int DefaultMaxSelectedMembers = 100;
    internal const int MaximumMaxSelectedMembers = 1000;

    internal static void ValidateMaxMembers(int maxMembers)
    {
        if (maxMembers is < 1 or > MaximumMaxSelectedMembers)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxMembers),
                $"Maximum selected members must be between 1 and {MaximumMaxSelectedMembers}.");
        }
    }

    internal static PivotSnapshotProperty<T> TryReadSnapshotProperty<T>(Func<T> read)
    {
        try
        {
            return new PivotSnapshotProperty<T>(true, read(), null);
        }
        catch (COMException ex) when (CreatedObjectFailure.CanReport(ex))
        {
            return new PivotSnapshotProperty<T>(false, default, ex.Message);
        }
        catch (NotSupportedException ex)
        {
            return new PivotSnapshotProperty<T>(false, default, ex.Message);
        }
    }

    internal static PivotFieldSelectionSnapshot CreateFieldSelection(
        string fieldName,
        PivotFieldArea area,
        int position,
        IEnumerable<string> selectedMembers,
        bool? showAll,
        bool? multipleSelectionEnabled,
        int maxMembers,
        bool fieldIdentityAvailable = true,
        bool selectionStateAvailable = true,
        string? selectionError = null,
        int? totalSelectedMemberCount = null)
    {
        ArgumentNullException.ThrowIfNull(selectedMembers);
        ArgumentOutOfRangeException.ThrowIfNegative(maxMembers);

        var result = new PivotFieldSelectionSnapshot
        {
            FieldName = fieldName,
            Area = area,
            Position = position,
            FieldIdentityAvailable = fieldIdentityAvailable,
            SelectionStateAvailable = selectionStateAvailable,
            SelectionError = selectionError,
            ShowAll = selectionStateAvailable ? showAll : null,
            MultipleSelectionEnabled = selectionStateAvailable ? multipleSelectionEnabled : null
        };

        if (!selectionStateAvailable || showAll is null || showAll.Value)
            return result;

        var semanticMembers = selectedMembers
            .Where(member => !string.IsNullOrWhiteSpace(member))
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToArray();

        if (totalSelectedMemberCount.HasValue && totalSelectedMemberCount.Value < semanticMembers.Length)
            throw new ArgumentOutOfRangeException(nameof(totalSelectedMemberCount));

        result.SelectedMemberCount = totalSelectedMemberCount ?? semanticMembers.Length;
        result.SelectedMemberUniqueNames = semanticMembers.Take(maxMembers).ToList();
        result.MembersTruncated = result.SelectedMemberCount > result.SelectedMemberUniqueNames.Count;
        return result;
    }
}

internal readonly record struct PivotSnapshotProperty<T>(
    bool Available,
    T? Value,
    string? ErrorMessage);
