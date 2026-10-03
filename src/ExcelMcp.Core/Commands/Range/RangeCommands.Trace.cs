using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeFormulaTraceResult TracePrecedents(IExcelBatch batch, string sheetName, string rangeAddress) =>
        TraceRelationships(batch, sheetName, rangeAddress, precedents: true);

    /// <inheritdoc />
    public RangeFormulaTraceResult TraceDependents(IExcelBatch batch, string sheetName, string rangeAddress) =>
        TraceRelationships(batch, sheetName, rangeAddress, precedents: false);

    private static RangeFormulaTraceResult TraceRelationships(
        IExcelBatch batch, string sheetName, string rangeAddress, bool precedents)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        return batch.Execute((context, ct) =>
        {
            Excel.Range? roots = null;
            Excel.Worksheet? sheet = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                roots = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (roots is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                sheet = roots.Worksheet;
                var result = new RangeFormulaTraceResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    RangeAddress = roots.Address,
                    Direction = precedents ? "precedents" : "dependents",
                    Action = precedents ? "trace-precedents" : "trace-dependents"
                };
                Dictionary<string, FormulaTraceNode> nodes = new(StringComparer.Ordinal);
                Queue<string> pending = new();
                HashSet<(string From, string To)> edges = [];
                RangeHelpers.VisitCells(roots, ct, cell => AddNode(cell, isRoot: true));
                while (pending.TryDequeue(out string? address))
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? cell = null;
                    Excel.Range? related = null;
                    Excel.Worksheet? relatedSheet = null;
                    try
                    {
                        cell = sheet.Range[address];
                        if (precedents && nodes[address].Formula is null)
                            continue;
                        try
                        {
                            related = precedents ? cell.DirectPrecedents : cell.DirectDependents;
                        }
                        catch (COMException exception) when (exception.HResult == unchecked((int)0x800A03EC))
                        {
                            result.Unresolved.Add(new FormulaTraceUnresolved(address,
                                "Native lookup returned no range; this can mean no relationships or unavailable " +
                                $"relationships. It is not classified as empty. Native diagnostic: {exception.Message}",
                                "0x800A03EC"));
                            continue;
                        }
                        if (related is null)
                            throw new InvalidOperationException("Excel returned a null native relationship range.");
                        relatedSheet = related.Worksheet;
                        if (!string.Equals(relatedSheet.Name, result.SheetName, StringComparison.Ordinal))
                            throw new InvalidOperationException(
                                "Excel returned a relationship outside the proven native worksheet boundary.");
                        RangeHelpers.VisitCells(related, ct, target =>
                        {
                            string targetAddress = target.Address;
                            edges.Add((address, targetAddress));
                            AddNode(target, isRoot: false);
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref relatedSheet);
                        ComUtilities.Release(ref related);
                        ComUtilities.Release(ref cell);
                    }
                }
                result.Nodes = nodes.Values.OrderBy(node => node.Row).ThenBy(node => node.Column).ToList();
                result.Edges = edges.OrderBy(edge => edge.From, StringComparer.Ordinal)
                    .ThenBy(edge => edge.To, StringComparer.Ordinal)
                    .Select(edge => new FormulaTraceEdge(edge.From, edge.To)).ToList();
                result.Cycles = FindTraceCycles(nodes.Keys, edges, ct);
                result.Coverage.NativeTraversalComplete = result.Unresolved.Count == 0;
                result.Success = true;
                return result;

                void AddNode(Excel.Range cell, bool isRoot)
                {
                    ct.ThrowIfCancellationRequested();
                    string address = cell.Address;
                    if (nodes.ContainsKey(address))
                        return;
                    object hasFormula = cell.HasFormula;
                    if (hasFormula is not bool formulaPresent)
                        throw new InvalidOperationException($"Excel returned an indeterminate formula state for {address}.");
                    string? formula = formulaPresent
                        ? Convert.ToString(ReadFormulas(context, cell), CultureInfo.InvariantCulture)
                        : null;
                    object? value = cell.Value2;
                    if (ExcelErrorMapper.TryGet(value, out _, out var mapped))
                        value = mapped.Name;
                    nodes.Add(address, new FormulaTraceNode(address, cell.Row, cell.Column, formula, value, isRoot));
                    pending.Enqueue(address);
                }
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref roots);
            }
        });
    }

    private static List<List<string>> FindTraceCycles(IEnumerable<string> addresses,
        HashSet<(string From, string To)> edges, CancellationToken ct)
    {
        var forward = addresses.ToDictionary(address => address, _ => new List<string>(), StringComparer.Ordinal);
        var reverse = addresses.ToDictionary(address => address, _ => new List<string>(), StringComparer.Ordinal);
        foreach (var edge in edges)
        {
            ct.ThrowIfCancellationRequested();
            forward[edge.From].Add(edge.To);
            reverse[edge.To].Add(edge.From);
        }
        HashSet<string> visited = new(StringComparer.Ordinal);
        List<string> finished = [];
        Stack<(string Address, bool Expanded)> stack = new();
        foreach (string address in forward.Keys)
        {
            ct.ThrowIfCancellationRequested();
            stack.Push((address, false));
            while (stack.TryPop(out var item))
            {
                ct.ThrowIfCancellationRequested();
                if (item.Expanded)
                {
                    finished.Add(item.Address);
                }
                else if (visited.Add(item.Address))
                {
                    stack.Push((item.Address, true));
                    foreach (string next in forward[item.Address])
                        stack.Push((next, false));
                }
            }
        }
        visited.Clear();
        List<List<string>> cycles = [];
        Stack<string> componentStack = new();
        for (int index = finished.Count - 1; index >= 0; index--)
        {
            ct.ThrowIfCancellationRequested();
            string start = finished[index];
            if (!visited.Add(start))
                continue;
            List<string> component = [];
            componentStack.Push(start);
            while (componentStack.TryPop(out string? address))
            {
                ct.ThrowIfCancellationRequested();
                component.Add(address);
                foreach (string previous in reverse[address])
                {
                    if (visited.Add(previous))
                        componentStack.Push(previous);
                }
            }
            if (component.Count > 1 || edges.Contains((start, start)))
            {
                component.Sort(StringComparer.Ordinal);
                cycles.Add(component);
            }
        }
        return cycles.OrderBy(component => component[0], StringComparer.Ordinal).ToList();
    }
}
