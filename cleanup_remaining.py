#!/usr/bin/env python3
"""
Final cleanup pass - remove remaining Task-related patterns that weren't caught.
"""

import os
import re

def cleanup_remaining_patterns():
    """Remove remaining Task/async patterns"""
    cli_commands_dir = r"d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands"
    
    command_files = [
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
    
    fixed_count = 0
    
    for filename in command_files:
        filepath = os.path.join(cli_commands_dir, filename)
        if not os.path.exists(filepath):
            continue
            
        with open(filepath, 'r') as f:
            content = f.read()
        
        original = content
        
        # Pattern 1: Remove "var task = " assignments and just use the result directly
        # var task = CommandHelper.WithBatchAsync(...); var result = task.GetAwaiter().GetResult();
        # becomes: var result = CommandHelper.WithBatchAsync(...);
        content = re.sub(
            r'var task = CommandHelper\.WithBatchAsync\(([^)]+(?:\([^)]*\))*[^)]*)\);[\s\n]+var result = task\.GetAwaiter\(\)\.GetResult\(\);',
            r'var result = CommandHelper.WithBatchAsync(\1);',
            content,
            flags=re.MULTILINE | re.DOTALL
        )
        
        # Pattern 2: Remove standing GetAwaiter().GetResult() calls
        content = re.sub(
            r'\.GetAwaiter\(\)\.GetResult\(\)',
            '',
            content
        )
        
        # Pattern 3: Remove remaining task. references when they're variables
        content = re.sub(
            r'var task = ([^;]+);[\s\n]*return task;',
            r'return \1;',
            content
        )
        
        # Pattern 4: Remove await keyword if it remains
        content = re.sub(
            r'\s+await\s+',
            ' ',
            content
        )
        
        if content != original:
            with open(filepath, 'w') as f:
                f.write(content)
            print(f"✅ Cleaned: {filename}")
            fixed_count += 1
    
    return fixed_count

if __name__ == '__main__':
    print("Removing remaining Task/async patterns...")
    num = cleanup_remaining_patterns()
    print(f"✅ Fixed {num} files")
