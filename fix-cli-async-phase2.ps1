# More comprehensive fix for async patterns
Get-ChildItem "d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands\*Commands.cs" | ForEach-Object {
    $file = $_
    $content = Get-Content $file.FullName -Raw
    $original = $content
    
    # Remove remaining .GetAwaiter().GetResult() calls (from incomplete Task.Run replacements)
    $content = $content -replace '\s*\.GetAwaiter\(\)\.GetResult\(\);', ';'
    
    # Remove async method declarations and replace with sync
    $content = $content -replace 'public async Task<int>', 'public int'
    $content = $content -replace 'private async Task<int>', 'private int'
    
    # Remove remaining async method bodies that have no actual async operations
    # This catches: private async Task<SomeResult> Method... 
    $content = $content -replace 'private async Task<(\w+)>', 'private $1'
    
    # Remove any remaining await keywords that might be stranded
    $content = $content -replace '\bawait\s+', ''
    
    # Fix any remaining issues with properties being called as methods (like .Code being mistaken for Code property)
    # This is harder without context, but we can detect some patterns
    
    if ($content -ne $original) {
        Write-Host "Further fixes: $($file.Name)"
        Set-Content $file.FullName -Value $content -Encoding UTF8
    }
}
