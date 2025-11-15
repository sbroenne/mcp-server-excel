#!/usr/bin/env python3
"""
Fix CLI async patterns to sync.
Replace all Task.Run(async () => { await using var batch = ... patterns
"""
import re
from pathlib import Path

def fix_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # First pass: Fix Task.Run(async () => { var batch = await ExcelSession.BeginBatch(...); return batch; })
    pattern_begin = r'var\s+task\s*=\s*Task\.Run\s*\(\s*async\s*\(\s*\)\s*=>\s*\{\s*var\s+batch\s*=\s*await\s+ExcelSession\.BeginBatch\s*\(\s*filePath\s*\)\s*;\s*return\s+batch\s*;\s*\}\s*\)\s*;\s*var\s+batch\s*=\s*task\.GetAwaiter\s*\(\s*\)\.GetResult\s*\(\s*\)\s*;'
    replacement_begin = r'var batch = ExcelSession.BeginBatch(filePath);'
    content = re.sub(pattern_begin, replacement_begin, content, flags=re.MULTILINE | re.DOTALL)
    
    # Pattern for Task.Run(async () => await batch.Dispose())
    pattern_dispose_task = r'var\s+disposeTask\s*=\s*Task\.Run\s*\(\s*async\s*\(\s*\)\s*=>\s*await\s+batch\.Dispose\s*\(\s*\)\s*\)\s*;\s*disposeTask\.GetAwaiter\s*\(\s*\)\.GetResult\s*\(\s*\)\s*;'
    replacement_dispose_task = r'batch.Dispose();'
    content = re.sub(pattern_dispose_task, replacement_dispose_task, content, flags=re.MULTILINE | re.DOTALL)
    
    # Pattern 3: using instead of await using
    content = content.replace('await using', 'using')
    
    # Pattern 4: BeginBatchAsync -> BeginBatch
    content = content.replace('ExcelSession.BeginBatchAsync', 'ExcelSession.BeginBatch')
    
    # Pattern 5: SaveAsync -> Save
    content = content.replace('.SaveAsync()', '.Save()')
    
    # Pattern 6: DisposeAsync -> Dispose (for batch)
    content = content.replace('.DisposeAsync()', '.Dispose()')
    
    # Pattern 7: Remove await keywords before ExcelSession.BeginBatch
    content = re.sub(r'await\s+ExcelSession\.BeginBatch', r'ExcelSession.BeginBatch', content)
    
    # Pattern 8: Remove remaining problematic Task.Run patterns
    pattern_task_generic = r'var\s+(\w+)\s*=\s*Task\.Run\s*\(\s*async\s*\(\s*\)\s*=>\s*\{([^}]+)\}\s*\)\.GetAwaiter\s*\(\s*\)\.GetResult\s*\(\s*\)\s*;'
    
    # For now, handle specific remaining patterns
    # Fix await batch.Dispose() inside Task.Run
    content = re.sub(r'await\s+batch\.Dispose\s*\(\s*\)', r'batch.Dispose()', content)
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

def main():
    cli_commands_dir = Path('src/ExcelMcp.CLI/Commands')
    
    # Files to fix
    files_to_fix = [
        'DataModelCommands.cs',
        'FileCommands.cs',
        'NamedRangeCommands.cs',
        'PivotTableCommands.cs',
        'QueryTableCommands.cs',
        'RangeCommands.cs',
        'TableCommands.cs',
        'ConditionalFormatCommands.cs',
        'PowerQueryCommands.cs',
        'SheetCommands.cs',
        'VbaCommands.cs',
        'BatchCommands.cs',
    ]
    
    fixed_count = 0
    for filename in files_to_fix:
        filepath = cli_commands_dir / filename
        if filepath.exists():
            if fix_file(filepath):
                print(f"Fixed: {filename}")
                fixed_count += 1
            else:
                print(f"No changes: {filename}")
        else:
            print(f"Not found: {filename}")
    
    print(f"\nTotal fixed: {fixed_count} files")

if __name__ == '__main__':
    main()

