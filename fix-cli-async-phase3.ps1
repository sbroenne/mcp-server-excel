# Phase 3: Fix remaining async/Task patterns properly
Get-ChildItem "d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands\*Commands.cs" | ForEach-Object {
    $file = $_
    $content = Get-Content $file.FullName -Raw
    $original = $content
    
    # Fix: var task = Task.Run(async () => { ... return _coreCommands.Method(...) }); result = task;
    # Replace with: var result = CommandHelper.WithBatch(args, filePath, false, batch => _coreCommands.Method(batch));
    # But this is complex - simpler approach: just unwrap the Task.Run
    
    # Pattern: "var task = Task.Run(async () => {\s*using var batch = ...\s*return (.*?)\s*});\s*result = task;"
    # Replace with: "using var batch = ...; var result = (get the return part without Task wrapper)"
    
    # More pragmatic: just remove the Task.Run wrapper and get the result properly
    # Pattern: var task = Task.Run(async () => { ... }); \n        result = task;
    # Replace with extraction of the inner logic
    
    # Simpler regex to remove Task.Run async wrappers where result assignment follows
    $content = [regex]::Replace($content,
        'var task = Task\.Run\(async \(\) =>\s*\{\s*using var batch = ExcelSession\.BeginBatch\(([^)]+)\);\s*return (.*?);?\s*\}\);\s*(\w+) = task;',
        'using var batch = ExcelSession.BeginBatch($1);
        $3 = $2;')
    
    # Also handle cases where result is declared and assigned differently
    $content = [regex]::Replace($content,
        'result = Task\.Run\(async \(\) =>\s*\{\s*using var batch = ExcelSession\.BeginBatch\(([^)]+)\);\s*return (.*?);?\s*\}\)\.GetAwaiter\(\)\.GetResult\(\);',
        'using var batch = ExcelSession.BeginBatch($1);
        result = $2;')
    
    if ($content -ne $original) {
        Write-Host "Phase 3 fixes: $($file.Name)"
        Set-Content $file.FullName -Value $content -Encoding UTF8
    }
}
