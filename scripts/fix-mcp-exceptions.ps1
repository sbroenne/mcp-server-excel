# fix-mcp-exceptions.ps1
# Adds error checking before JsonSerializer.Serialize returns

param(
    [switch]$DryRun = $false
)

$toolsPath = "D:\source\mcp-server-excel\src\ExcelMcp.McpServer\Tools"
$files = Get-ChildItem $toolsPath -Filter "*.cs"

$totalFixed = 0
$fileResults = @()

foreach ($file in $files) {
    $content = Get-Content $file.FullName -Raw
    $originalContent = $content
    $fixCount = 0
    
    # Pattern: Find JsonSerializer.Serialize(result, ...) returns without error check
    # Look for pattern where there's no "if (!result.Success" before the return
    
    $lines = Get-Content $file.FullName
    $newLines = @()
    $i = 0
    
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        
        # Check if this line is a return JsonSerializer.Serialize(result
        if ($line -match '^\s+return JsonSerializer\.Serialize\(result') {
            # Look back up to 5 lines to see if there's already error checking
            $hasErrorCheck = $false
            $lookBack = [Math]::Min(5, $i)
            
            for ($j = 1; $j -le $lookBack; $j++) {
                $prevLine = $lines[$i - $j]
                if ($prevLine -match 'if\s*\(\s*!result\.Success' -or 
                    $prevLine -match 'throw.*McpException') {
                    $hasErrorCheck = $true
                    break
                }
            }
            
            if (-not $hasErrorCheck) {
                # Need to add error check - extract indent
                if ($line -match '^(\s+)return') {
                    $indent = $matches[1]
                    
                    # Try to extract action name and parameter from method context
                    # Look back for method signature
                    $methodLine = ""
                    for ($k = $i; $k -ge 0; $k--) {
                        if ($lines[$k] -match 'private static async Task<string>\s+(\w+)') {
                            $methodLine = $lines[$k]
                            break
                        }
                    }
                    
                    # Add error check lines before return
                    $newLines += ""
                    $newLines += "$indent// Check for errors and throw McpException"
                    $newLines += "${indent}if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))"
                    $newLines += "$indent{"
                    $newLines += "$indent    throw new ModelContextProtocol.McpException(`$`"Operation failed: {result.ErrorMessage}`");  
                    $newLines += "$indent}"
                    $newLines += ""
                    
                    $fixCount++
                }
            }
        }
        
        $newLines += $line
        $i++
    }
    
    if ($fixCount -gt 0) {
        $fileResults += [PSCustomObject]@{
            File = $file.Name
            FixCount = $fixCount
        }
        
        if (-not $DryRun) {
            Set-Content $file.FullName ($newLines -join "`r`n") -NoNewline
        }
        
        $totalFixed += $fixCount
    }
}

# Display results
Write-Output "`n═══════════════════════════════════════════"
Write-Output "  MCP Exception Handling Fix Results"
Write-Output "═══════════════════════════════════════════`n"

if ($DryRun) {
    Write-Output "🔍 DRY RUN MODE - No changes made`n"
}

if ($fileResults.Count -eq 0) {
    Write-Output "✅ No fixes needed - all tools already throw McpException!"
} else {
    $fileResults | Format-Table -AutoSize
    Write-Output "`n📊 Total fixes: $totalFixed methods across $($fileResults.Count) files"
    
    if ($DryRun) {
        Write-Output "`n💡 Run without -DryRun to apply changes"
    } else {
        Write-Output "`n✅ Changes applied successfully!"
    }
}
