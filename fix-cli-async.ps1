# Fix async patterns in CLI Commands
Get-ChildItem "d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands\*Commands.cs" | ForEach-Object {
    $file = $_
    $content = Get-Content $file.FullName -Raw
    $original = $content
    
    # Replace SaveAsync() with Save()
    $content = $content -replace 'await batch\.SaveAsync\(\)', 'batch.Save()'
    
    # Replace DisposeAsync() with Dispose()
    $content = $content -replace 'await batch\.DisposeAsync\(\)', 'batch.Dispose()'
    $content = $content -replace '\$?.* \= Task\.Run\(async \(\) => await batch\.DisposeAsync\(\)\);\s*\$?.*.GetAwaiter\(\)\.GetResult\(\);', 'batch.Dispose();'
    
    # Replace _coreCommands.XxxAsync with _coreCommands.Xxx
    $content = $content -replace '_coreCommands\.(\w+)Async\(', '_coreCommands.$1('
    
    # Count changes
    if ($content -ne $original) {
        Write-Host "Fixing: $($file.Name)"
        Set-Content $file.FullName -Value $content -Encoding UTF8
    }
}
