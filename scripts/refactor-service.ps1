# Refactor ExcelMcpService.cs: Replace hand-written handlers with generated dispatch
$file = "d:\source\mcp-server-excel\src\ExcelMcp.Service\ExcelMcpService.cs"
$lines = Get-Content $file -Encoding utf8

# Find the line containing "// === SHEET COMMANDS ===" -- start of DELETE section
$startPattern = "// === SHEET COMMANDS ==="
$startIdx = -1
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match [regex]::Escape($startPattern)) {
        $startIdx = $i
        break
    }
}
if ($startIdx -eq -1) { throw "Could not find start marker: $startPattern" }

# Find the line containing "private Task<ServiceResponse> WithSessionAsync" -- start of KEEP section
$endPattern = "private Task<ServiceResponse> WithSessionAsync"
$endIdx = -1
for ($i = $startIdx; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match [regex]::Escape($endPattern)) {
        $endIdx = $i
        break
    }
}
if ($endIdx -eq -1) { throw "Could not find end marker: $endPattern" }

Write-Host "Deleting lines $($startIdx+1) to $($endIdx) (0-based: $startIdx to $($endIdx-1))"
Write-Host "Total lines to delete: $($endIdx - $startIdx)"

# Build new dispatch methods block
$newBlock = @'

    // === GENERATED DISPATCH ===
    // All command routing uses ServiceRegistry.*.DispatchToCore() generated methods.
    // See ServiceRegistry.*.Dispatch.g.cs for the generated code.

    private delegate bool TryParseDelegate<TAction>(string action, out TAction result);

    private static ServiceResponse WrapResult(string? dispatchResult)
    {
        return dispatchResult == null
            ? new ServiceResponse { Success = true }
            : new ServiceResponse { Success = true, Result = dispatchResult };
    }

    private async Task<ServiceResponse> DispatchSimpleAsync<TAction>(
        string actionString, ServiceRequest request,
        TryParseDelegate<TAction> tryParse,
        Func<TAction, IExcelBatch, string?> dispatch) where TAction : struct
    {
        if (!tryParse(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return await WithSessionAsync(request.SessionId, batch => WrapResult(dispatch(action, batch)));
    }

    private async Task<ServiceResponse> DispatchSheetAsync(string actionString, ServiceRequest request)
    {
        if (ServiceRegistry.Sheet.TryParseAction(actionString, out var sheetAction))
        {
            // CopyToFile/MoveToFile are atomic operations without session
            if (sheetAction is SheetAction.CopyToFile or SheetAction.MoveToFile)
            {
                try
                {
                    return WrapResult(ServiceRegistry.Sheet.DispatchToCore(
                        _sheetCommands, sheetAction, null!, request.Args));
                }
                catch (Exception ex)
                {
                    return new ServiceResponse { Success = false, ErrorMessage = ex.Message };
                }
            }

            return await WithSessionAsync(request.SessionId, batch =>
                WrapResult(ServiceRegistry.Sheet.DispatchToCore(_sheetCommands, sheetAction, batch, request.Args)));
        }

        if (ServiceRegistry.SheetStyle.TryParseAction(actionString, out var styleAction))
        {
            return await WithSessionAsync(request.SessionId, batch =>
                WrapResult(ServiceRegistry.SheetStyle.DispatchToCore(_sheetCommands, styleAction, batch, request.Args)));
        }

        return new ServiceResponse { Success = false, ErrorMessage = $"Unknown sheet action: {actionString}" };
    }

    private async Task<ServiceResponse> DispatchRangeAsync(string actionString, ServiceRequest request)
    {
        return await WithSessionAsync(request.SessionId, batch =>
        {
            if (ServiceRegistry.Range.TryParseAction(actionString, out var ra))
                return WrapResult(ServiceRegistry.Range.DispatchToCore(_rangeCommands, ra, batch, request.Args));
            if (ServiceRegistry.RangeEdit.TryParseAction(actionString, out var rea))
                return WrapResult(ServiceRegistry.RangeEdit.DispatchToCore(_rangeCommands, rea, batch, request.Args));
            if (ServiceRegistry.RangeFormat.TryParseAction(actionString, out var rfa))
                return WrapResult(ServiceRegistry.RangeFormat.DispatchToCore(_rangeCommands, rfa, batch, request.Args));
            if (ServiceRegistry.RangeLink.TryParseAction(actionString, out var rla))
                return WrapResult(ServiceRegistry.RangeLink.DispatchToCore(_rangeCommands, rla, batch, request.Args));
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown range action: {actionString}" };
        });
    }

    private async Task<ServiceResponse> DispatchTableAsync(string actionString, ServiceRequest request)
    {
        return await WithSessionAsync(request.SessionId, batch =>
        {
            if (ServiceRegistry.Table.TryParseAction(actionString, out var ta))
                return WrapResult(ServiceRegistry.Table.DispatchToCore(_tableCommands, ta, batch, request.Args));
            if (ServiceRegistry.TableColumn.TryParseAction(actionString, out var tca))
                return WrapResult(ServiceRegistry.TableColumn.DispatchToCore(_tableCommands, tca, batch, request.Args));
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown table action: {actionString}" };
        });
    }

'@

# Build new file content
$before = $lines[0..($startIdx - 1)]
$after = $lines[$endIdx..($lines.Count - 1)]
$newContent = $before + $newBlock.Split("`n") + $after

# Write file
$newContent | Set-Content $file -Encoding utf8
Write-Host "Refactored. New file has $($newContent.Count) lines (was $($lines.Count) lines)"
