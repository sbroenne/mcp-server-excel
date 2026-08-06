// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Service.Safety;

internal static class WorkbookSemanticInspector
{
    private const int MaximumTargetCells = 10_000;
    private const int MaximumWorkbookCells = 20_000;

    public static SemanticSnapshot Capture(
        IExcelBatch batch,
        CommandSafetyDescriptor descriptor,
        string? argsJson)
    {
        var target = ResolveTarget(argsJson);

        return batch.Execute((context, cancellationToken) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            var workbookParts = new List<string>();
            var hasExactRangeTarget = descriptor.ScopeResolver == "range" &&
                !string.IsNullOrWhiteSpace(target.RangeAddress);
            var hasExactCopyTarget = descriptor.ScopeResolver == "rangeCopy" &&
                !string.IsNullOrWhiteSpace(target.SourceSheetName) &&
                !string.IsNullOrWhiteSpace(target.SourceRangeAddress) &&
                !string.IsNullOrWhiteSpace(target.TargetSheetName) &&
                !string.IsNullOrWhiteSpace(target.TargetRangeAddress);

            if (hasExactCopyTarget)
            {
                CaptureWorkbookCollections(
                    context.Book,
                    workbookParts,
                    includeWorksheetIdentities: true);
                return CaptureRangeCopy(
                    batch,
                    context.Book,
                    descriptor,
                    target,
                    workbookParts);
            }

            if (hasExactRangeTarget)
            {
                // Range-targeted reviews still bind to the workbook's structural
                // collections. A table/name/chart/query change outside the target
                // range must invalidate the review before the write is dispatched.
                // Worksheet identities are collected in the same pass as their
                // collections to avoid a duplicate cross-process COM traversal.
                CaptureWorkbookCollections(
                    context.Book,
                    workbookParts,
                    includeWorksheetIdentities: true);
                return CaptureRange(
                    batch,
                    context.Book,
                    descriptor,
                    target.SheetName,
                    target.RangeAddress!,
                    workbookParts,
                    workbookBounded: true);
            }

            var bounded = true;
            CaptureWorkbookHeader(context.Book, workbookParts, ref bounded);
            CaptureWorkbookCollections(context.Book, workbookParts);
            var scope = ResolveScope(descriptor.ScopeResolver, target);
            var fingerprint = SafetyFingerprint.Hash(descriptor.Command, string.Join('|', workbookParts));
            return new SemanticSnapshot(
                fingerprint,
                fingerprint,
                scope,
                descriptor.VerificationLevel,
                bounded,
                [],
                0);
        });
    }

    /// <summary>
    /// Captures only the evidence needed to verify an already-authorized exact-range
    /// mutation. The complete workbook structure is always revalidated immediately
    /// before dispatch by <see cref="Capture"/>; repeating that authoritative scan
    /// after the write cannot make the authorization safer and would dominate the
    /// cost of checking the affected cells.
    /// </summary>
    public static SemanticSnapshot CapturePostMutation(
        IExcelBatch batch,
        CommandSafetyDescriptor descriptor,
        string? argsJson)
    {
        var target = ResolveTarget(argsJson);
        var hasExactCopyTarget = descriptor.ScopeResolver == "rangeCopy" &&
            !string.IsNullOrWhiteSpace(target.SourceSheetName) &&
            !string.IsNullOrWhiteSpace(target.SourceRangeAddress) &&
            !string.IsNullOrWhiteSpace(target.TargetSheetName) &&
            !string.IsNullOrWhiteSpace(target.TargetRangeAddress);
        if (hasExactCopyTarget)
        {
            return batch.Execute((context, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                var source = CaptureRange(
                    batch,
                    context.Book,
                    descriptor,
                    target.SourceSheetName,
                    target.SourceRangeAddress!,
                    [],
                    workbookBounded: true);
                var targetAddress = ResolveCopyDestinationAddress(
                    context.Book,
                    target.TargetSheetName!,
                    target.TargetRangeAddress!,
                    source.RowCount,
                    source.ColumnCount);
                return CaptureRange(
                    batch,
                    context.Book,
                    descriptor,
                    target.TargetSheetName,
                    targetAddress,
                    [],
                    source.IsBounded);
            });
        }

        var hasExactRangeTarget = descriptor.ScopeResolver == "range" &&
            !string.IsNullOrWhiteSpace(target.RangeAddress);
        if (!hasExactRangeTarget)
        {
            return Capture(batch, descriptor, argsJson);
        }

        return batch.Execute((context, cancellationToken) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            return CaptureRange(
                batch,
                context.Book,
                descriptor,
                target.SheetName,
                target.RangeAddress!,
                [],
                workbookBounded: true);
        });
    }

    /// <summary>
    /// Resolves only the identifiers which can be safely derived from a command payload.
    /// Keeping this separate from COM inspection makes the review scope deterministic and testable.
    /// </summary>
    internal static SemanticInspectionTarget ResolveTarget(string? argsJson)
    {
        if (string.IsNullOrWhiteSpace(argsJson))
        {
            return new SemanticInspectionTarget();
        }

        try
        {
            using var document = JsonDocument.Parse(argsJson);
            var root = document.RootElement;
            return new SemanticInspectionTarget(
                FindString(root, "sheetName", "sheet"),
                FindString(root, "rangeAddress", "range"),
                FindString(root, "tableName"),
                FindString(root, "chartName"),
                FindString(root, "pivotTableName"),
                FindString(root, "queryName"),
                FindString(root, "connectionName"),
                FindString(root, "name"),
                FindString(root, "sourceSheet", "sourceSheetName"),
                FindString(root, "sourceRange", "sourceAddress"),
                FindString(root, "targetSheet", "targetSheetName"),
                FindString(root, "targetRange", "targetAddress"));
        }
        catch (JsonException)
        {
            return new SemanticInspectionTarget();
        }
    }

    internal static SafetyScope ResolveScope(string scopeResolver, SemanticInspectionTarget target)
    {
        var sheets = target.SheetName is null ? new List<string>() : [target.SheetName];
        var objects = new List<string>();

        switch (scopeResolver)
        {
            case "worksheet" when target.SheetName is not null:
                objects.Add($"worksheet:{target.SheetName}");
                break;
            case "table" when target.TableName is not null:
                objects.Add($"table:{target.TableName}");
                break;
            case "chart" when target.ChartName is not null:
                objects.Add($"chart:{target.ChartName}");
                break;
            case "pivotTable" when target.PivotTableName is not null:
                objects.Add($"pivotTable:{target.PivotTableName}");
                break;
            case "externalObject":
                if (target.QueryName is not null) objects.Add($"powerQuery:{target.QueryName}");
                if (target.ConnectionName is not null) objects.Add($"connection:{target.ConnectionName}");
                break;
            case "workbook" when target.Name is not null:
                objects.Add($"name:{target.Name}");
                break;
        }

        return sheets.Count == 0 && objects.Count == 0
            ? SafetyScope.Workbook
            : new SafetyScope(sheets, [], objects);
    }

    public static VerificationReceipt Compare(SemanticSnapshot before, SemanticSnapshot after)
    {
        var changedCells = CountChangedCells(before.CellHashes, after.CellHashes);
        var bounded = before.IsBounded && after.IsBounded;
        var hasExactRangeEvidence = before.VerificationLevel == "rangeSemantic" &&
            bounded &&
            before.CellCount > 0 &&
            before.CellHashes.Count == before.CellCount &&
            after.CellHashes.Count == after.CellCount &&
            before.CellCount == after.CellCount &&
            before.Scope.Ranges.Count == 1 &&
            HasSameScope(before.Scope, after.Scope);

        var status = before.VerificationLevel == "notVerified"
            ? "notVerified"
            : hasExactRangeEvidence
                ? "verified"
                : "partiallyVerified";

        var limitation = status switch
        {
            "verified" => null,
            "notVerified" => "This operation has no reliable post-mutation semantic inspection surface.",
            _ => "Verification was limited to a bounded semantic fingerprint of the inspected scope."
        };

        return new VerificationReceipt(
            status,
            before.Scope,
            changedCells,
            before.VerificationFingerprint,
            after.VerificationFingerprint,
            limitation);
    }

    private static SemanticSnapshot CaptureRange(
        IExcelBatch batch,
        dynamic workbook,
        CommandSafetyDescriptor descriptor,
        string? sheetName,
        string rangeAddress,
        IReadOnlyList<string> workbookParts,
        bool workbookBounded)
    {
        dynamic? worksheet = null;
        dynamic? rangeWorksheet = null;
        dynamic? definedName = null;
        dynamic? application = null;
        dynamic? range = null;
        dynamic? rows = null;
        dynamic? columns = null;

        try
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                worksheet = workbook.Worksheets.Item[sheetName];
                range = worksheet.Range[rangeAddress];
            }
            else
            {
                try
                {
                    definedName = workbook.Names.Item(rangeAddress);
                    range = definedName.RefersToRange;
                }
                catch (Exception ex) when (IsInspectableFailure(ex))
                {
                    ComUtilities.Release(ref definedName);
                    application = workbook.Application;
                    range = application.Range[rangeAddress];
                }
            }

            rangeWorksheet = range.Worksheet;
            var resolvedSheetName = Convert.ToString(rangeWorksheet.Name, CultureInfo.InvariantCulture) ?? sheetName ?? "Workbook";
            rows = range.Rows;
            columns = range.Columns;
            var rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
            var columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            var cellCount = checked(rowCount * columnCount);
            var bounded = workbookBounded && cellCount <= MaximumTargetCells;

            var cellHashes = new List<string>();
            if (cellCount <= MaximumTargetCells)
            {
                var values = Flatten(range.Value2, cellCount);
                var formulas = Flatten(FormulaCompatibility.Read(batch, range), cellCount);
                for (var index = 0; index < cellCount; index++)
                {
                    var value = index < values.Count ? values[index] : "missing";
                    var formula = index < formulas.Count ? formulas[index] : "missing";
                    cellHashes.Add(SafetyFingerprint.Hash(value, formula));
                }
            }

            var resolvedAddress = Convert.ToString(range.Address, CultureInfo.InvariantCulture) ?? rangeAddress;
            var scopeAddress = $"{resolvedSheetName}!{resolvedAddress}";
            var verificationFingerprint = SafetyFingerprint.Hash(
                descriptor.Command,
                scopeAddress,
                rowCount.ToString(CultureInfo.InvariantCulture),
                columnCount.ToString(CultureInfo.InvariantCulture),
                string.Join('|', cellHashes));
            var fingerprint = SafetyFingerprint.Hash(
                verificationFingerprint,
                string.Join('|', workbookParts));

            return new SemanticSnapshot(
                fingerprint,
                verificationFingerprint,
                new SafetyScope([resolvedSheetName], [scopeAddress], []),
                descriptor.VerificationLevel,
                bounded,
                cellHashes,
                cellCount,
                rowCount,
                columnCount);
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref rangeWorksheet);
            ComUtilities.Release(ref definedName);
            ComUtilities.Release(ref application);
            ComUtilities.Release(ref worksheet);
        }
    }

    private static SemanticSnapshot CaptureRangeCopy(
        IExcelBatch batch,
        dynamic workbook,
        CommandSafetyDescriptor descriptor,
        SemanticInspectionTarget target,
        IReadOnlyList<string> workbookParts)
    {
        var source = CaptureRange(
            batch,
            (object)workbook,
            descriptor,
            target.SourceSheetName,
            target.SourceRangeAddress!,
            workbookParts,
            workbookBounded: true);
        var targetAddress = ResolveCopyDestinationAddress(
            (object)workbook,
            target.TargetSheetName!,
            target.TargetRangeAddress!,
            source.RowCount,
            source.ColumnCount);
        var destination = CaptureRange(
            batch,
            (object)workbook,
            descriptor,
            target.TargetSheetName,
            targetAddress,
            Array.Empty<string>(),
            source.IsBounded);

        return destination with
        {
            Fingerprint = SafetyFingerprint.Hash(source.Fingerprint, destination.Fingerprint),
            IsBounded = source.IsBounded && destination.IsBounded
        };
    }

    private static string ResolveCopyDestinationAddress(
        object workbookObject,
        string targetSheetName,
        string targetRangeAddress,
        int rowCount,
        int columnCount)
    {
        dynamic workbook = workbookObject;
        dynamic? worksheet = null;
        dynamic? destination = null;
        dynamic? resized = null;
        try
        {
            worksheet = workbook.Worksheets.Item[targetSheetName];
            destination = worksheet.Range[targetRangeAddress];
            resized = destination.Resize[rowCount, columnCount];
            return Convert.ToString(resized.Address, CultureInfo.InvariantCulture) ?? targetRangeAddress;
        }
        finally
        {
            ComUtilities.Release(ref resized);
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref worksheet);
        }
    }

    private static void CaptureWorkbookHeader(dynamic workbook, List<string> parts, ref bool bounded)
    {
        dynamic? worksheets = null;
        var remainingCells = MaximumWorkbookCells;

        try
        {
            worksheets = workbook.Worksheets;
            var count = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            parts.Add($"worksheets:{count}");

            for (var index = 1; index <= count; index++)
            {
                dynamic? worksheet = null;
                dynamic? usedRange = null;
                dynamic? rows = null;
                dynamic? columns = null;
                try
                {
                    worksheet = worksheets.Item[index];
                    usedRange = worksheet.UsedRange;
                    rows = usedRange.Rows;
                    columns = usedRange.Columns;
                    var name = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture) ?? $"Sheet{index}";
                    var visibility = Convert.ToString(worksheet.Visible, CultureInfo.InvariantCulture) ?? string.Empty;
                    var address = Convert.ToString(usedRange.Address, CultureInfo.InvariantCulture) ?? string.Empty;
                    var rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
                    var columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
                    var cellCount = checked(rowCount * columnCount);
                    parts.Add($"sheet:{index}:{name}:{visibility}:{address}:{rowCount}:{columnCount}");

                    if (cellCount <= remainingCells)
                    {
                        parts.Add(SafetyFingerprint.Hash(
                            string.Join('|', Flatten(usedRange.Value2, cellCount)),
                            string.Join('|', Flatten(usedRange.Formula, cellCount))));
                        remainingCells -= cellCount;
                    }
                    else
                    {
                        bounded = false;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref columns);
                    ComUtilities.Release(ref rows);
                    ComUtilities.Release(ref usedRange);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }

    // These collections describe workbook structure which is not represented by UsedRange.
    // They are deliberately captured for every resolver: an intervening structural change must
    // invalidate a review even when the command itself has a narrower target.
    private static void CaptureWorkbookCollections(
        dynamic workbook,
        List<string> parts,
        bool includeWorksheetIdentities = false)
    {
        CaptureWorkbookNames(workbook, parts);
        CaptureWorkbookQueries(workbook, parts);
        CaptureWorkbookConnections(workbook, parts);
        CaptureWorkbookRelationships(workbook, parts);

        dynamic? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            var count = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            if (includeWorksheetIdentities)
            {
                parts.Add($"worksheets:{count}");
            }

            for (var index = 1; index <= count; index++)
            {
                dynamic? worksheet = null;
                try
                {
                    worksheet = worksheets.Item[index];
                    var sheetName = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture) ?? $"Sheet{index}";
                    if (includeWorksheetIdentities)
                    {
                        var visibility = Convert.ToString(worksheet.Visible, CultureInfo.InvariantCulture) ?? string.Empty;
                        parts.Add($"sheet:{index}:{sheetName}:{visibility}");
                    }

                    CaptureWorksheetTables(worksheet, sheetName, parts);
                    CaptureWorksheetCharts(worksheet, sheetName, parts);
                    CaptureWorksheetPivotTables(worksheet, sheetName, parts);
                }
                finally
                {
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"worksheets-collections:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }

    private static void CaptureWorkbookNames(dynamic workbook, List<string> parts)
    {
        dynamic? names = null;
        try
        {
            names = workbook.Names;
            CaptureCollection((object)names, "name", parts, item =>
                $"{Stable(item.Name)}:{Stable(item.RefersTo)}:{Stable(item.Visible)}");
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"names:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref names);
        }
    }

    private static void CaptureWorkbookQueries(dynamic workbook, List<string> parts)
    {
        dynamic? queries = null;
        try
        {
            queries = workbook.Queries;
            CaptureCollection((object)queries, "powerQuery", parts, item =>
                $"{Stable(item.Name)}:{Stable(item.Formula)}");
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"powerQueries:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref queries);
        }
    }

    private static void CaptureWorkbookConnections(dynamic workbook, List<string> parts)
    {
        dynamic? connections = null;
        try
        {
            connections = workbook.Connections;
            CaptureCollection((object)connections, "connection", parts, item =>
                $"{Stable(item.Name)}:{Stable(item.Type)}:{Stable(item.Description)}");
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"connections:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    private static void CaptureWorkbookRelationships(dynamic workbook, List<string> parts)
    {
        dynamic? model = null;
        dynamic? relationships = null;
        try
        {
            model = workbook.Model;
            relationships = model.ModelRelationships;
            CaptureCollection((object)relationships, "relationship", parts, DescribeRelationship);
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            // The data model is optional and unavailable in many Excel editions.
            parts.Add($"relationships:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref relationships);
            ComUtilities.Release(ref model);
        }
    }

    private static string DescribeRelationship(dynamic relationship)
    {
        dynamic? foreignKey = null;
        dynamic? primaryKey = null;
        dynamic? foreignTable = null;
        dynamic? primaryTable = null;
        try
        {
            foreignKey = relationship.ForeignKeyColumn;
            primaryKey = relationship.PrimaryKeyColumn;
            foreignTable = foreignKey.Parent;
            primaryTable = primaryKey.Parent;
            return $"{Stable(foreignTable.Name)}:{Stable(foreignKey.Name)}:{Stable(primaryTable.Name)}:{Stable(primaryKey.Name)}:{Stable(relationship.Active)}";
        }
        finally
        {
            ComUtilities.Release(ref primaryTable);
            ComUtilities.Release(ref foreignTable);
            ComUtilities.Release(ref primaryKey);
            ComUtilities.Release(ref foreignKey);
        }
    }

    private static void CaptureWorksheetTables(dynamic worksheet, string sheetName, List<string> parts)
    {
        dynamic? tables = null;
        try
        {
            tables = worksheet.ListObjects;
            CaptureCollection((object)tables, $"table:{sheetName}", parts, item =>
            {
                dynamic? range = null;
                try
                {
                    range = item.Range;
                    return $"{Stable(item.Name)}:{Stable(range.Address)}:{Stable(item.ShowTotals)}:{Stable(item.TableStyle)}";
                }
                finally
                {
                    ComUtilities.Release(ref range);
                }
            });
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"tables:{sheetName}:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref tables);
        }
    }

    private static void CaptureWorksheetCharts(dynamic worksheet, string sheetName, List<string> parts)
    {
        dynamic? charts = null;
        try
        {
            charts = worksheet.ChartObjects();
            CaptureCollection((object)charts, $"chart:{sheetName}", parts, item =>
            {
                dynamic? chart = null;
                try
                {
                    chart = item.Chart;
                    return $"{Stable(item.Name)}:{Stable(item.Left)}:{Stable(item.Top)}:{Stable(item.Width)}:{Stable(item.Height)}:{Stable(chart.ChartType)}:{Stable(chart.HasTitle)}";
                }
                finally
                {
                    ComUtilities.Release(ref chart);
                }
            });
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"charts:{sheetName}:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref charts);
        }
    }

    private static void CaptureWorksheetPivotTables(dynamic worksheet, string sheetName, List<string> parts)
    {
        dynamic? pivots = null;
        try
        {
            pivots = worksheet.PivotTables();
            CaptureCollection((object)pivots, $"pivotTable:{sheetName}", parts, item =>
            {
                dynamic? range = null;
                try
                {
                    range = item.TableRange2;
                    return $"{Stable(item.Name)}:{Stable(range.Address)}:{Stable(item.ManualUpdate)}";
                }
                finally
                {
                    ComUtilities.Release(ref range);
                }
            });
        }
        catch (Exception ex) when (IsInspectableFailure(ex))
        {
            parts.Add($"pivotTables:{sheetName}:unavailable:{ex.GetType().Name}");
        }
        finally
        {
            ComUtilities.Release(ref pivots);
        }
    }

    private static void CaptureCollection(object collection, string label, List<string> parts, Func<dynamic, string> describe)
    {
        dynamic dynamicCollection = collection;
        var count = Convert.ToInt32(dynamicCollection.Count, CultureInfo.InvariantCulture);
        parts.Add($"{label}:count:{count}");
        for (var index = 1; index <= count; index++)
        {
            dynamic? item = null;
            try
            {
                item = dynamicCollection.Item(index);
                parts.Add($"{label}:item:{index}:{describe(item)}");
            }
            finally
            {
                ComUtilities.Release(ref item);
            }
        }
    }

    private static string Stable(object? value) => ToStableScalar(value);

    internal static bool IsInspectableFailure(Exception exception) =>
        exception is not OutOfMemoryException and not StackOverflowException and not OperationCanceledException &&
        !IsFatalExcelDisconnect(exception);

    private static bool IsFatalExcelDisconnect(Exception exception)
    {
        for (var current = exception; current is not null; current = current.InnerException)
        {
            if (current is COMException comException &&
                (comException.HResult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
                 comException.HResult == ResiliencePipelines.RPC_E_CALL_FAILED ||
                 comException.HResult == ResiliencePipelines.RPC_E_DISCONNECTED))
            {
                return true;
            }
        }

        return false;
    }

    private static List<string> Flatten(object? value, int expectedCount)
    {
        var values = new List<string>(Math.Max(0, expectedCount));
        if (value is Array array)
        {
            if (array.Rank == 2)
            {
                for (var row = array.GetLowerBound(0); row <= array.GetUpperBound(0); row++)
                {
                    for (var column = array.GetLowerBound(1); column <= array.GetUpperBound(1); column++)
                    {
                        values.Add(ToStableScalar(array.GetValue(row, column)));
                    }
                }
            }
            else
            {
                foreach (var item in array)
                {
                    values.Add(ToStableScalar(item));
                }
            }
        }
        else
        {
            values.Add(ToStableScalar(value));
        }

        return values;
    }

    private static string ToStableScalar(object? value)
    {
        if (value is null) return "null";
        if (value is DateTime dateTime) return $"datetime:{dateTime.ToUniversalTime():O}";
        if (value is double number) return $"number:{number.ToString("R", CultureInfo.InvariantCulture)}";
        if (value is float single) return $"number:{single.ToString("R", CultureInfo.InvariantCulture)}";
        if (value is IFormattable formattable) return $"{value.GetType().Name}:{formattable.ToString(null, CultureInfo.InvariantCulture)}";
        return $"{value.GetType().Name}:{Convert.ToString(value, CultureInfo.InvariantCulture)}";
    }

    private static int CountChangedCells(IReadOnlyList<string> before, IReadOnlyList<string> after)
    {
        var count = Math.Max(before.Count, after.Count);
        var changed = 0;
        for (var index = 0; index < count; index++)
        {
            var beforeHash = index < before.Count ? before[index] : null;
            var afterHash = index < after.Count ? after[index] : null;
            if (!string.Equals(beforeHash, afterHash, StringComparison.Ordinal))
            {
                changed++;
            }
        }

        return changed;
    }

    private static bool HasSameScope(SafetyScope before, SafetyScope after) =>
        before.Sheets.SequenceEqual(after.Sheets, StringComparer.Ordinal) &&
        before.Ranges.SequenceEqual(after.Ranges, StringComparer.Ordinal) &&
        before.Objects.SequenceEqual(after.Objects, StringComparer.Ordinal);

    private static string? FindString(JsonElement root, params string[] names)
    {
        if (root.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        foreach (var property in root.EnumerateObject())
        {
            if (property.Value.ValueKind == JsonValueKind.String &&
                names.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
            {
                return property.Value.GetString();
            }
        }

        return null;
    }
}

internal sealed record SemanticInspectionTarget(
    string? SheetName = null,
    string? RangeAddress = null,
    string? TableName = null,
    string? ChartName = null,
    string? PivotTableName = null,
    string? QueryName = null,
    string? ConnectionName = null,
    string? Name = null,
    string? SourceSheetName = null,
    string? SourceRangeAddress = null,
    string? TargetSheetName = null,
    string? TargetRangeAddress = null);
