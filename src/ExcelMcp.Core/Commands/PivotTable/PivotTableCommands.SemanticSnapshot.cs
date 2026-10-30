using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using ComUtilities = Sbroenne.ExcelMcp.ComInterop.ComUtilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

public partial class PivotTableCommands
{
    private static Excel.PivotTable FindPivotTable(Excel.Workbook workbook, string sheetName, string pivotTableName)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.PivotTables? pivotTables = null;
        try
        {
            sheets = (Excel.Sheets)workbook.Worksheets;
            sheet = (Excel.Worksheet)sheets[sheetName];
            pivotTables = (Excel.PivotTables)sheet.PivotTables();
            return pivotTables.Item(pivotTableName);
        }
        finally
        {
            ComUtilities.Release(ref pivotTables);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }

    private static List<PivotFieldSelectionSnapshot> GetFieldSelectionSnapshots(
        bool isOlap,
        Excel.CubeFields? cubeFields,
        Excel.PivotFields? pivotFields,
        string pivotTableName,
        string workbookPath,
        int maxSelectedMembers,
        CancellationToken cancellationToken,
        IExcelBatchWindowControl? windowControl)
    {
        var selections = new List<PivotFieldSelectionSnapshot>();
        var returnedMembers = new SortedSet<SnapshotMember>(SnapshotMemberComparer.Instance);

        if (isOlap)
        {
            if (cubeFields is null)
                throw new InvalidOperationException("Excel identified an OLAP PivotTable but did not provide its CubeFields collection.");

            for (int index = 1; index <= cubeFields.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.CubeField? cubeField = null;
                try
                {
                    cubeField = cubeFields[index];
                    var selection = ReadOlapFieldSelection(
                        cubeField, pivotTableName, workbookPath, cancellationToken, maxSelectedMembers, windowControl);
                    if (selection is null)
                        continue;

                    selections.Add(selection);
                    if (selection.SelectionStateAvailable && selection.ShowAll == false)
                    {
                        foreach (string member in selection.SelectedMemberUniqueNames)
                        {
                            returnedMembers.Add(new SnapshotMember(
                                selection.FieldName, selection.Area, selection.Position, member));
                            if (returnedMembers.Count > maxSelectedMembers)
                                returnedMembers.Remove(returnedMembers.Max!);
                        }
                    }

                    selection.SelectedMemberUniqueNames.Clear();
                }
                finally
                {
                    ComUtilities.Release(ref cubeField);
                }
            }
        }
        else
        {
            if (pivotFields is null)
                throw new InvalidOperationException("Excel did not provide the PivotFields collection for this PivotTable.");

            for (int index = 1; index <= pivotFields.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.PivotField? field = null;
                try
                {
                    field = pivotFields.Item(index);
                    var selection = ReadRegularFieldSelection(
                        field, pivotTableName, workbookPath, cancellationToken, maxSelectedMembers);
                    if (selection is null)
                        continue;

                    selections.Add(selection);
                    if (selection.SelectionStateAvailable && selection.ShowAll == false)
                    {
                        foreach (string member in selection.SelectedMemberUniqueNames)
                        {
                            returnedMembers.Add(new SnapshotMember(
                                selection.FieldName, selection.Area, selection.Position, member));
                            if (returnedMembers.Count > maxSelectedMembers)
                                returnedMembers.Remove(returnedMembers.Max!);
                        }
                    }

                    selection.SelectedMemberUniqueNames.Clear();
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }
        }

        foreach (var selection in selections)
        {
            selection.SelectedMemberUniqueNames = returnedMembers
                .Where(member => member.FieldName == selection.FieldName &&
                    member.Area == selection.Area &&
                    member.Position == selection.Position)
                .Select(member => member.MemberName)
                .ToList();

            if (selection.SelectedMemberCount.HasValue)
                selection.MembersTruncated = selection.SelectedMemberCount.Value > selection.SelectedMemberUniqueNames.Count;
        }

        return selections
            .OrderBy(selection => selection.FieldName, StringComparer.Ordinal)
            .ThenBy(selection => selection.Area)
            .ThenBy(selection => selection.Position)
            .ToList();
    }

    internal static PivotFieldSelectionSnapshot? ReadRegularFieldSelection(
        Excel.PivotField field,
        string pivotTableName,
        string workbookPath,
        CancellationToken cancellationToken,
        int maxSelectedMembers)
    {
        var rawSourceName = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() => (object?)field.SourceName);
        if (rawSourceName.Available && PivotTableHelpers.IsValuesPlaceholderField(field, rawSourceName.Value))
            return null;

        var sourceName = new PivotSnapshotProperty<string>(
            rawSourceName.Available, rawSourceName.Value?.ToString(), rawSourceName.ErrorMessage);
        var nativeName = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => field.Name?.ToString());
        string fieldName = !string.IsNullOrWhiteSpace(sourceName.Value)
            ? sourceName.Value!
            : nativeName.Value ?? string.Empty;
        bool identityAvailable = sourceName.Available &&
            (!string.IsNullOrWhiteSpace(sourceName.Value) ||
                (nativeName.Available && !string.IsNullOrWhiteSpace(nativeName.Value)));
        var caption = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => field.Caption?.ToString());
        var orientation = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => Convert.ToInt32(field.Orientation, System.Globalization.CultureInfo.InvariantCulture));

        var result = new PivotFieldSelectionSnapshot
        {
            FieldName = fieldName,
            Caption = caption.Available ? caption.Value : null,
            FieldIdentityAvailable = identityAvailable,
            IdentityError = identityAvailable ? null : sourceName.ErrorMessage ?? nativeName.ErrorMessage,
            Area = PivotFieldArea.Hidden,
            LayoutAvailable = false,
            LayoutError = orientation.ErrorMessage,
            SelectionStateAvailable = false,
            SelectionError = "The field identity or placement could not be read.",
            FiltersError = "The field identity or placement could not be read."
        };

        if (!orientation.Available)
            return result;

        if (!Enum.IsDefined(typeof(PivotFieldArea), orientation.Value))
        {
            result.LayoutError = $"Excel returned an unknown field orientation value: {orientation.Value}.";
            return result;
        }

        result.Area = (PivotFieldArea)orientation.Value;
        if (result.Area is not (PivotFieldArea.Row or PivotFieldArea.Column or PivotFieldArea.Filter))
            return null;

        if (!identityAvailable)
            return result;

        var position = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() => Convert.ToInt32(
            (object)field.Position, System.Globalization.CultureInfo.InvariantCulture));
        if (!position.Available)
        {
            result.LayoutError = position.ErrorMessage;
            return result;
        }

        result.Position = position.Value;
        result.LayoutAvailable = true;
        var memberState = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => ReadRegularMemberState(field, result, maxSelectedMembers, cancellationToken));
        if (!memberState.Available || memberState.Value is null)
        {
            result.SelectionError = memberState.ErrorMessage;
        }
        else
        {
            result.SelectionStateAvailable = true;
            result.ShowAll = memberState.Value.ShowAll;
            result.MultipleSelectionEnabled = memberState.Value.MultipleSelectionEnabled;
            result.SelectedMemberCount = memberState.Value.SelectedMemberCount;
            result.SelectedMemberUniqueNames = memberState.Value.MemberNames;
            result.MembersTruncated = memberState.Value.SelectedMemberCount.HasValue &&
                memberState.Value.SelectedMemberCount.Value >
                memberState.Value.MemberNames.Count;
        }

        var filters = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => ReadPivotFilters(field, pivotTableName, workbookPath, cancellationToken));
        result.FiltersAvailable = filters.Available;
        result.FiltersError = filters.ErrorMessage;
        if (filters.Available && filters.Value is not null)
        {
            result.Filters = filters.Value.Filters.Select(filter => new PivotSnapshotFilter
            {
                Type = filter.Type,
                Value1 = filter.Value1,
                Value2 = filter.Value2,
                DataFieldName = filter.DataFieldName
            }).ToList();
        }

        return result;
    }

    internal static PivotFieldSelectionSnapshot? ReadOlapFieldSelection(
        Excel.CubeField cubeField,
        string pivotTableName,
        string workbookPath,
        CancellationToken cancellationToken,
        int maxSelectedMembers,
        IExcelBatchWindowControl? windowControl)
    {
        var name = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() => cubeField.Name);
        var caption = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() => cubeField.Caption);
        var orientation = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => Convert.ToInt32(cubeField.Orientation, System.Globalization.CultureInfo.InvariantCulture));
        string fieldName = name.Value ?? string.Empty;
        var result = new PivotFieldSelectionSnapshot
        {
            FieldName = fieldName,
            Caption = caption.Available ? caption.Value : null,
            FieldIdentityAvailable = name.Available && !string.IsNullOrWhiteSpace(fieldName),
            IdentityError = name.ErrorMessage,
            LayoutAvailable = false,
            LayoutError = orientation.ErrorMessage ?? name.ErrorMessage,
            SelectionStateAvailable = false,
            SelectionError = "The hierarchy identity or placement could not be read.",
            FiltersError = "The hierarchy identity or placement could not be read."
        };

        if (!orientation.Available)
            return result;
        if (!Enum.IsDefined(typeof(PivotFieldArea), orientation.Value))
        {
            result.LayoutError = $"Excel returned an unknown hierarchy orientation value: {orientation.Value}.";
            return result;
        }

        result.Area = (PivotFieldArea)orientation.Value;
        if (result.Area is not (PivotFieldArea.Row or PivotFieldArea.Column or PivotFieldArea.Filter))
            return null;
        if (!result.FieldIdentityAvailable)
            return result;

        var position = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
            () => Convert.ToInt32(cubeField.Position, System.Globalization.CultureInfo.InvariantCulture));
        if (!position.Available)
        {
            result.LayoutError = position.ErrorMessage;
            return result;
        }

        result.Position = position.Value;
        result.LayoutAvailable = true;
        Excel.PivotFields? levels = null;
        var levelFields = new List<Excel.PivotField>();
        try
        {
            var levelsResult = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
                () => (Excel.PivotFields)cubeField.PivotFields);
            if (!levelsResult.Available || levelsResult.Value is null)
            {
                result.SelectionError = levelsResult.ErrorMessage ??
                    "Excel did not expose PivotFields for this placed hierarchy.";
                result.FiltersError = result.SelectionError;
                return result;
            }

            levels = levelsResult.Value;
            var countResult = PivotTableSnapshotHelpers.TryReadSnapshotProperty(() => levels.Count);
            if (!countResult.Available)
            {
                result.SelectionError = countResult.ErrorMessage;
                result.FiltersError = countResult.ErrorMessage;
                return result;
            }

            if (countResult.Value < 1)
            {
                result.SelectionError = "Excel did not expose a PivotField for this placed hierarchy.";
                result.FiltersError = result.SelectionError;
                return result;
            }

            for (int levelIndex = 1; levelIndex <= countResult.Value; levelIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var levelResult = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
                    () => (Excel.PivotField)levels.Item(levelIndex));
                if (!levelResult.Available || levelResult.Value is null)
                {
                    result.SelectionError = levelResult.ErrorMessage ??
                        $"Excel did not expose level {levelIndex} of this placed hierarchy.";
                    result.FiltersError = result.SelectionError;
                    return result;
                }

                levelFields.Add(levelResult.Value);
            }

            var memberState = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
                () => ReadOlapMemberState(cubeField, levelFields, result, maxSelectedMembers, cancellationToken, windowControl));
            if (!memberState.Available || memberState.Value is null)
            {
                result.SelectionError = memberState.ErrorMessage;
            }
            else
            {
                result.SelectionStateAvailable = true;
                result.ShowAll = memberState.Value.ShowAll;
                result.MultipleSelectionEnabled = memberState.Value.MultipleSelectionEnabled;
                result.SelectedMemberCount = memberState.Value.SelectedMemberCount;
                result.SelectedMemberUniqueNames = memberState.Value.MemberNames;
                result.MembersTruncated = memberState.Value.SelectedMemberCount.HasValue &&
                    memberState.Value.SelectedMemberCount.Value >
                    memberState.Value.MemberNames.Count;
            }

            var filters = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
                () => ReadOlapLevelFilters(levelFields, pivotTableName, workbookPath, cancellationToken));
            result.FiltersAvailable = filters.Available;
            result.FiltersError = filters.ErrorMessage;
            if (filters.Available && filters.Value is not null)
                result.Filters = filters.Value;
        }
        finally
        {
            for (int index = levelFields.Count - 1; index >= 0; index--)
            {
                var levelField = levelFields[index];
                ComUtilities.Release(ref levelField);
            }

            ComUtilities.Release(ref levels);
        }

        return result;
    }

    private static List<PivotSnapshotFilter> ReadOlapLevelFilters(
        List<Excel.PivotField> levelFields,
        string pivotTableName,
        string workbookPath,
        CancellationToken cancellationToken)
    {
        var filters = new List<PivotSnapshotFilter>();
        foreach (var levelField in levelFields)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int orientation = Convert.ToInt32(levelField.Orientation, System.Globalization.CultureInfo.InvariantCulture);
            if (orientation == (int)PivotFieldArea.Hidden)
                continue;

            string levelName = levelField.Name;
            var levelFilters = ReadPivotFilters(levelField, pivotTableName, workbookPath, cancellationToken, isOlap: true);
            var unidentifiableValueFilter = levelFilters.Filters.FirstOrDefault(filter =>
                (filter.Type.StartsWith("Value", StringComparison.Ordinal) ||
                    filter.Type.StartsWith("Top", StringComparison.Ordinal) ||
                    filter.Type.StartsWith("Bottom", StringComparison.Ordinal)) &&
                filter.DataFieldName is null);
            if (unidentifiableValueFilter is not null)
                throw new NotSupportedException(
                    $"Excel did not expose the value field for OLAP filter '{unidentifiableValueFilter.Type}' on level '{levelName}'; the semantic snapshot cannot be complete.");
            filters.AddRange(levelFilters.Filters.Select(filter => new PivotSnapshotFilter
            {
                LevelUniqueName = levelName,
                Type = filter.Type,
                Value1 = filter.Value1,
                Value2 = filter.Value2,
                DataFieldName = filter.DataFieldName
            }));
        }

        return filters;
    }

    private static SnapshotMemberState ReadRegularMemberState(
        Excel.PivotField field,
        PivotFieldSelectionSnapshot selection,
        int maxMembers,
        CancellationToken cancellationToken)
    {
        bool isFilter = selection.Area == PivotFieldArea.Filter;
        bool? multiple = isFilter ? field.EnableMultiplePageItems : null;
        if (field.AllItemsVisible)
            return new SnapshotMemberState(true, multiple, null, []);

        if (isFilter && multiple == false)
        {
            Excel.PivotItem? currentPage = null;
            try
            {
                currentPage = (Excel.PivotItem)field.CurrentPage;
                string currentPageName = currentPage.Name ?? string.Empty;
                bool showAll = IsAllRegularPivotMember(currentPageName);
                var names = new BoundedMemberNames(maxMembers);
                if (!showAll)
                    names.Add(currentPageName);
                return new SnapshotMemberState(showAll, false, showAll ? null : names.Count, names.Names);
            }
            finally
            {
                ComUtilities.Release(ref currentPage);
            }
        }

        // Excel reports every member as visible for report-filter fields, so only row and
        // column fields can read the visible members without walking every member.
        return isFilter
            ? ReadAllRegularMembers(field, multiple, maxMembers, cancellationToken)
            : ReadVisibleRegularMembers(field, maxMembers, cancellationToken);
    }

    private static SnapshotMemberState ReadVisibleRegularMembers(
        Excel.PivotField field,
        int maxMembers,
        CancellationToken cancellationToken)
    {
        Excel.PivotItems? visibleItems = null;
        try
        {
            visibleItems = (Excel.PivotItems)field.VisibleItems[Type.Missing];
            int visibleCount = visibleItems.Count;
            if (visibleCount == 0)
                throw new NotSupportedException("Excel returned no visible items, so the field's selection state is unknown.");

            var names = new BoundedMemberNames(maxMembers);
            for (int index = 1; index <= visibleCount; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.PivotItem? item = null;
                try
                {
                    item = (Excel.PivotItem)field.VisibleItems[index];
                    names.Add(item.Name);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            return new SnapshotMemberState(false, null, visibleCount, names.Names);
        }
        finally
        {
            ComUtilities.Release(ref visibleItems);
        }
    }

    private static SnapshotMemberState ReadAllRegularMembers(
        Excel.PivotField field,
        bool? multiple,
        int maxMembers,
        CancellationToken cancellationToken)
    {
        Excel.PivotItems? items = null;
        try
        {
            items = (Excel.PivotItems)field.PivotItems();
            int count = items.Count;
            if (count == 0)
                throw new NotSupportedException("Excel returned no items, so the field's selection state is unknown.");

            var names = new BoundedMemberNames(maxMembers);
            int visibleCount = 0;
            for (int index = 1; index <= count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.PivotItem? item = null;
                try
                {
                    item = items.Item(index);
                    if (item.Visible)
                    {
                        visibleCount++;
                        names.Add(item.Name);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            bool showAll = visibleCount == count;
            return new SnapshotMemberState(
                showAll,
                multiple,
                showAll ? null : visibleCount,
                names.Names);
        }
        finally
        {
            ComUtilities.Release(ref items);
        }
    }

    private static SnapshotMemberState ReadOlapMemberState(
        Excel.CubeField cubeField,
        List<Excel.PivotField> levelFields,
        PivotFieldSelectionSnapshot selection,
        int maxMembers,
        CancellationToken cancellationToken,
        IExcelBatchWindowControl? windowControl)
    {
        if (selection.Area == PivotFieldArea.Filter)
        {
            bool multiple = cubeField.EnableMultiplePageItems;
            if (!multiple)
            {
                Excel.PivotField levelField = levelFields[0];
                string currentPage = levelField.CurrentPageName ?? string.Empty;
                bool showAll = PivotTableHelpers.IsAllOlapReportFilterMember(
                    cubeField, currentPage, cancellationToken, windowControl);
                var names = new BoundedMemberNames(maxMembers);
                if (!showAll)
                    names.Add(currentPage);
                return new SnapshotMemberState(showAll, false, showAll ? null : names.Count, names.Names);
            }
        }

        if (cubeField.AllItemsVisible)
            return new SnapshotMemberState(true, selection.Area == PivotFieldArea.Filter ? true : null, null, []);

        // Each level keeps its own selection; unique names carry the level path, so the
        // union across levels is the hierarchy's exact selection.
        var selectedNames = new BoundedMemberNames(maxMembers);
        foreach (var levelField in levelFields)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (levelField.VisibleItemsList is not Array items)
                throw new NotSupportedException("Excel did not return a member list for an OLAP hierarchy level.");

            foreach (object? item in items)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string? memberName = item?.ToString();
                if (!string.IsNullOrWhiteSpace(memberName))
                    selectedNames.Add(memberName);
            }
        }

        if (selectedNames.Count == 0)
            throw new NotSupportedException("Excel reported a filtered OLAP hierarchy but returned no selected member unique names.");

        return new SnapshotMemberState(
            false,
            selection.Area == PivotFieldArea.Filter ? true : null,
            selectedNames.Count,
            selectedNames.Names);
    }

    private static bool IsAllRegularPivotMember(string memberName) =>
        string.IsNullOrWhiteSpace(memberName) ||
        string.Equals(memberName, "(All)", StringComparison.OrdinalIgnoreCase) ||
        memberName.EndsWith(".[All]", StringComparison.OrdinalIgnoreCase);

    private sealed record SnapshotMemberState(
        bool ShowAll,
        bool? MultipleSelectionEnabled,
        int? SelectedMemberCount,
        List<string> MemberNames);

    private sealed class BoundedMemberNames(int maxMembers)
    {
        private readonly SortedSet<string> _names = new(StringComparer.Ordinal);

        public int Count { get; private set; }
        public List<string> Names => _names.ToList();

        public void Add(string memberName)
        {
            Count++;
            _names.Add(memberName);
            if (_names.Count > maxMembers)
                _names.Remove(_names.Max!);
        }
    }

    private readonly record struct SnapshotMember(
        string FieldName,
        PivotFieldArea Area,
        int Position,
        string MemberName);

    private sealed class SnapshotMemberComparer : IComparer<SnapshotMember>
    {
        public static SnapshotMemberComparer Instance { get; } = new();

        public int Compare(SnapshotMember x, SnapshotMember y)
        {
            int comparison = StringComparer.Ordinal.Compare(x.FieldName, y.FieldName);
            if (comparison != 0)
                return comparison;
            comparison = x.Area.CompareTo(y.Area);
            if (comparison != 0)
                return comparison;
            comparison = x.Position.CompareTo(y.Position);
            return comparison != 0
                ? comparison
                : StringComparer.Ordinal.Compare(x.MemberName, y.MemberName);
        }
    }
}
