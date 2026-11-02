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
    public RangeFormulaTraceResult TracePrecedents(IExcelBatch batch, string sheetName, string rangeAddress,
        int maxDepth = 3, int maxNodes = 50, int maxPreviewCharacters = 256) =>
        TraceRelationships(batch, sheetName, rangeAddress, true, maxDepth, maxNodes, maxPreviewCharacters);

    /// <inheritdoc />
    public RangeFormulaTraceResult TraceDependents(IExcelBatch batch, string sheetName, string rangeAddress,
        int maxDepth = 3, int maxNodes = 50, int maxPreviewCharacters = 256) =>
        TraceRelationships(batch, sheetName, rangeAddress, false, maxDepth, maxNodes, maxPreviewCharacters);

    private static RangeFormulaTraceResult TraceRelationships(
        IExcelBatch batch, string sheetName, string rangeAddress, bool precedents,
        int maxDepth, int maxNodes, int maxPreviewCharacters)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        if (maxDepth is < 0 or > 20)
            throw new ArgumentOutOfRangeException(nameof(maxDepth), "maxDepth must be between 0 and 20.");
        if (maxNodes is < 1 or > 1000)
            throw new ArgumentOutOfRangeException(nameof(maxNodes), "maxNodes must be between 1 and 1000.");
        if (maxPreviewCharacters is < 1 or > 4096)
            throw new ArgumentOutOfRangeException(nameof(maxPreviewCharacters), "maxPreviewCharacters must be between 1 and 4096.");
        return batch.Execute(ReadTrace);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "Native relationship lookup errors are reported as unresolved with their exact diagnostic and HRESULT, never classified as an empty graph; traversal completeness is false.")]
        RangeFormulaTraceResult ReadTrace(ExcelContext context, CancellationToken ct)
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
                    Action = precedents ? "trace-precedents" : "trace-dependents",
                    Limits = new FormulaTraceLimits(maxDepth, maxNodes, maxPreviewCharacters)
                };
                Dictionary<string, FormulaTraceNode> nodes = new(StringComparer.Ordinal);
                Queue<string> pending = new();
                HashSet<(string From, string To)> edges = [];
                RangeHelpers.VisitCellsWhile(roots, ct, cell =>
                {
                    if (nodes.ContainsKey(cell.Address))
                        return true;
                    if (nodes.Count == maxNodes)
                    {
                        result.Truncation.StartingCellsOmitted = true;
                        result.Truncation.NodeLimitReached = true;
                        result.Boundaries.Add(new FormulaTraceBoundary(result.RangeAddress, "max-nodes"));
                        return false;
                    }
                    AddNode(cell, isRoot: true, depth: 0);
                    return true;
                });
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
                        RangeHelpers.VisitCellsWhile(related, ct, target =>
                        {
                            string targetAddress = target.Address;
                            if (!nodes.ContainsKey(targetAddress))
                            {
                                bool depthExceeded = nodes[address].Depth == maxDepth;
                                bool nodesExceeded = nodes.Count == maxNodes;
                                if (depthExceeded || nodesExceeded)
                                {
                                    if (depthExceeded)
                                    {
                                        result.Truncation.DepthLimitReached = true;
                                        result.Boundaries.Add(new FormulaTraceBoundary(address, "max-depth"));
                                    }
                                    if (nodesExceeded)
                                    {
                                        result.Truncation.NodeLimitReached = true;
                                        result.Boundaries.Add(new FormulaTraceBoundary(address, "max-nodes"));
                                    }
                                    return false;
                                }
                                AddNode(target, isRoot: false, depth: nodes[address].Depth + 1);
                            }
                            edges.Add((address, targetAddress));
                            return true;
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
                result.Coverage.NativeTraversalComplete = result.Unresolved.Count == 0 &&
                    !result.Truncation.DepthLimitReached && !result.Truncation.NodeLimitReached;
                result.Success = true;
                return result;

                void AddNode(Excel.Range cell, bool isRoot, int depth)
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
                    if (ExcelErrorMapper.TryGet(value, out _, out var mapped) && IsNativeError(context, cell, 1, 1))
                        value = mapped.Name;
                    string? formulaPreview = LimitTraceText(formula, maxPreviewCharacters);
                    object? valuePreview = value is string text ? LimitTraceText(text, maxPreviewCharacters) : value;
                    bool formulaTruncated = formulaPreview != formula;
                    bool valueTruncated = !Equals(valuePreview, value);
                    result.Truncation.PreviewTextTruncated |= formulaTruncated || valueTruncated;
                    nodes.Add(address, new FormulaTraceNode(address, cell.Row, cell.Column, formulaPreview,
                        valuePreview, isRoot, depth, formulaTruncated, valueTruncated));
                    pending.Enqueue(address);
                }
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref roots);
            }
        }
    }

    private static string? LimitTraceText(string? text, int limit)
    {
        if (text is null || text.Length <= limit)
            return text;
        if (char.IsHighSurrogate(text[limit - 1]) && char.IsLowSurrogate(text[limit]))
            limit--;
        return text[..limit];
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
