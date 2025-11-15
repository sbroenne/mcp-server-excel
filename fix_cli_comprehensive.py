#!/usr/bin/env python3
"""
Fix CLI async patterns to sync using safe, tested patterns.
This script applies only the patterns that work correctly.
"""
import re
from pathlib import Path

def replace_in_file(file_path, pattern, replacement):
    """Safely replace a pattern in a file"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original = content
    # Use verbose mode for clarity
    content = re.sub(pattern, replacement, content, flags=re.MULTILINE | re.DOTALL | re.VERBOSE)
    
    if content != original:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

def fix_command_helper():
    """Fix CommandHelper.cs"""
    file = Path('src/ExcelMcp.CLI/Commands/CommandHelper.cs')
    if not file.exists():
        return False
    
    # Change method signature from async to sync
    pattern = r'public\s+static\s+async\s+Task<T>\s+WithBatchAsync'
    replacement = r'public static T WithBatch'
    replace_in_file(file, pattern, replacement)
    
    # Replace the parameter type
    pattern = r'Func<IExcelBatch,\s*Task<T>>\s+action'
    replacement = r'Func<IExcelBatch, T> action'
    replace_in_file(file, pattern, replacement)
    
    # Replace entire first branch (session case)
    pattern = r'var\s+task\s*=\s*Task\.Run\s*\(\s*async\s*\(\s*\)\s*=>\s*await\s+action\s*\(\s*batch\s*\)\s*\)\s*;\s*return\s+task\.GetAwaiter\s*\(\s*\)\.GetResult\s*\(\s*\)\s*;'
    replacement = r'return action(batch);'
    replace_in_file(file, pattern, replacement)
    
    # Replace entire second branch (session-of-one case)
    pattern = r'''var\s+task\s*=\s*Task\.Run\s*\(\s*async\s*\(\s*\)\s*=>\s*\{
\s*await\s+using\s+var\s+batch\s*=\s*await\s+ExcelSession\.BeginBatchAsync\s*\(\s*filePath\s*\)\s*;
\s*var\s+result\s*=\s*await\s+action\s*\(\s*batch\s*\)\s*;

\s*if\s*\(\s*save\s*\)\s*\{
\s*await\s+batch\.SaveAsync\s*\(\s*\)\s*;
\s*\}

\s*return\s+result\s*;
\s*\}\s*\)\s*;
\s*return\s+task\.GetAwaiter\s*\(\s*\)\.GetResult\s*\(\s*\)\s*;'''
    replacement = r'''using var batch = ExcelSession.BeginBatch(filePath);
            var result = action(batch);

            if (save)
            {
                batch.Save();
            }

            return result;'''
    replace_in_file(file, pattern, replacement)
    
    return True

def fix_simple_patterns():
    """Fix simple string replacements across all files"""
    cli_commands_dir = Path('src/ExcelMcp.CLI/Commands')
    
    files_to_fix = [
        'CommandHelper.cs',
        'PowerQueryCommands.cs',
        'SheetCommands.cs',
        'VbaCommands.cs',
        'BatchCommands.cs',
        'ConnectionCommands.cs',
        'DataModelCommands.cs',
        'FileCommands.cs',
        'NamedRangeCommands.cs',
        'PivotTableCommands.cs',
        'QueryTableCommands.cs',
        'RangeCommands.cs',
        'TableCommands.cs',
        'ConditionalFormatCommands.cs',
    ]
    
    for filename in files_to_fix:
        filepath = cli_commands_dir / filename
        if not filepath.exists():
            continue
        
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original = content
        
        # Simple string replacements first
        content = content.replace('await using var batch = await ExcelSession.BeginBatchAsync(filePath)',
                                  'using var batch = ExcelSession.BeginBatch(filePath)')
        content = content.replace('.SaveAsync()', '.Save()')
        content = content.replace('.DisposeAsync()', '.Dispose()')
        content = content.replace('await batch.Save()', 'batch.Save()')
        content = content.replace('await batch.Dispose()', 'batch.Dispose()')
        
        # Remove Task.Run and GetAwaiter().GetResult() patterns
        # This is tricky - we need to be very careful
        # Match: var <name> = Task.Run(async () => { ... }); var <name2> = <name>.GetAwaiter().GetResult();
        # For simple cases with single statements
        
        pattern = r'var\s+(\w+)\s*=\s*Task\.Run\s*\(\s*async\s*\(\s*\)\s*=>\s*\{\s*return\s+await\s+_coreCommands\.(\w+)\s*\(\s*batch\s*(?:,\s*([^)]+))?\s*\)\s*;\s*\}\s*\)\s*;\s*var\s+result\s*=\s*\1\.GetAwaiter\s*\(\s*\)\.GetResult\s*\(\s*\)\s*;'
        
        # This would be: var task = Task.Run(async () => { return await _coreCommands.Method(batch, args); }); var result = task.GetAwaiter().GetResult();
        # To: var result = _coreCommands.Method(batch, args);
        def replace_task_pattern(match):
            task_var = match.group(1)
            method = match.group(2)
            args = match.group(3) or ''
            if args:
                return f'var result = _coreCommands.{method}(batch, {args});'
            else:
                return f'var result = _coreCommands.{method}(batch);'
        
        content = re.sub(pattern, replace_task_pattern, content, flags=re.MULTILINE | re.DOTALL)
        
        if content != original:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"Fixed: {filename}")
        else:
            print(f"No changes: {filename}")

def main():
    print("Fixing CLI async patterns to sync...")
    print()
    
    # First fix CommandHelper
    print("Step 1: Fixing CommandHelper.cs...")
    fix_command_helper()
    
    # Then fix simple patterns in all files
    print("\nStep 2: Fixing simple patterns in all command files...")
    fix_simple_patterns()
    
    print("\nDone!")

if __name__ == '__main__':
    main()
