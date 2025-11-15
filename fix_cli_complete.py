#!/usr/bin/env python3
"""
Complete CLI async-to-sync refactoring for ExcelMcp.
Handles all patterns needed to convert async CLI to synchronous patterns.
"""

import os
import re

def fix_command_helper():
    """Fix CommandHelper.cs to use sync Core API"""
    filepath = r"d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands\CommandHelper.cs"
    with open(filepath, 'r') as f:
        content = f.read()
    
    original = content
    
    # Pattern 1: Fix the method signature - change Func<IExcelBatch, Task<T>> to Func<IExcelBatch, T>
    content = re.sub(
        r'Func<IExcelBatch,\s*Task<T>>\s+action',
        'Func<IExcelBatch, T> action',
        content
    )
    
    # Pattern 2: Fix the async lambda in existing session branch to just call action
    content = re.sub(
        r'var task = Task\.Run\(async \(\) => await action\(batch\)\);',
        'var result = action(batch);',
        content
    )
    content = re.sub(
        r'return task\.GetAwaiter\(\)\.GetResult\(\);',
        'return result;',
        content,
        count=1  # Only first occurrence after the session-of-one is fixed
    )
    
    # Pattern 3: Fix session-of-one to not use async/Task.Run
    # Replace the entire async block with sync version
    old_session_block = r'''var task = Task\.Run\(async \(\) =>\s*\{[\s\S]*?await using var batch = await ExcelSession\.BeginBatchAsync\(filePath\);[\s\S]*?var result = await action\(batch\);[\s\S]*?if \(save\)[\s\S]*?\{[\s\S]*?await batch\.SaveAsync\(\);[\s\S]*?\}[\s\S]*?return result;[\s\S]*?\}\);[\s\S]*?return task\.GetAwaiter\(\)\.GetResult\(\);'''
    new_session_block = '''using var batch = ExcelSession.BeginBatch(filePath);
            var result = action(batch);

            if (save)
            {
                batch.Save();
            }

            return result;'''
    
    content = re.sub(old_session_block, new_session_block, content)
    
    if content != original:
        with open(filepath, 'w') as f:
            f.write(content)
        print("✅ Fixed: CommandHelper.cs")
        return True
    return False

def fix_cli_command_files():
    """Fix all CLI command files to be synchronous"""
    cli_commands_dir = r"d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands"
    
    # Files to fix (all command implementations)
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
        
        # Pattern 1: Remove 'async' from method signatures that have no await
        # Match: public int MethodName(string[] args) { ... } that's marked async Task<int>
        content = re.sub(
            r'(\s+)public async Task<int> (\w+)\(string\[\] args\)',
            r'\1public int \2(string[] args)',
            content
        )
        
        # Pattern 2: Replace CommandHelper.WithBatchAsync with CommandHelper.WithBatch calls
        # Change _coreCommands.MethodAsync to _coreCommands.Method
        content = re.sub(
            r'CommandHelper\.WithBatchAsync\(',
            'CommandHelper.WithBatchAsync(',  # Keep for now, we'll fix the name later
            content
        )
        
        # Pattern 3: Remove 'await' keyword before method calls on _coreCommands
        content = re.sub(
            r'await\s+_coreCommands\.(\w+Async)\(',
            r'_coreCommands.\1(',
            content
        )
        
        # Pattern 4: Convert MethodAsync to Method (remove Async suffix)
        # Only for _coreCommands method calls
        content = re.sub(
            r'_coreCommands\.(\w+)Async\(',
            r'_coreCommands.\1(',
            content
        )
        
        # Pattern 5: Fix parameter lambda functions - remove 'async' and 'await'
        # Pattern: (batch) => await _coreCommands.MethodAsync(...) becomes (batch) => _coreCommands.Method(...)
        content = re.sub(
            r'\(batch\)\s*=>\s*await\s+',
            '(batch) => ',
            content
        )
        
        # Pattern 6: Remove 'await' from 'await using' (convert to 'using')
        content = re.sub(
            r'await using',
            'using',
            content
        )
        
        # Pattern 7: Fix ExcelSession calls - BeginBatchAsync to BeginBatch
        content = re.sub(
            r'ExcelSession\.BeginBatchAsync\(',
            'ExcelSession.BeginBatch(',
            content
        )
        
        # Pattern 8: Fix SaveAsync to Save
        content = re.sub(
            r'batch\.SaveAsync\(',
            'batch.Save(',
            content
        )
        
        # Pattern 9: Fix DisposeAsync to Dispose (if using explicit dispose)
        content = re.sub(
            r'batch\.DisposeAsync\(\)',
            'batch.Dispose()',
            content
        )
        
        # Pattern 10: Remove any remaining 'async' keywords with no 'await'
        # Match lines that are async methods with no awaits in their scope
        # This is complex, so we'll do a simpler pass: remove 'async' from all method signatures
        content = re.sub(
            r'(\s+)private async Task<(\w+)> (\w+)\(',
            r'\1private \2 \3(',
            content
        )
        
        content = re.sub(
            r'(\s+)private async void (\w+)\(',
            r'\1private void \2(',
            content
        )
        
        # Pattern 11: Remove trailing await keywords that are lonely
        content = re.sub(
            r'return await (\w+);',
            r'return \1;',
            content
        )
        
        if content != original:
            with open(filepath, 'w') as f:
                f.write(content)
            print(f"✅ Fixed: {filename}")
            fixed_count += 1
    
    return fixed_count

def fix_interface_files():
    """Fix interface signatures to be synchronous"""
    interfaces = [
        r"d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands\IPowerQueryCommands.cs",
        r"d:\source\mcp-server-excel\src\ExcelMcp.CLI\Commands\IVbaCommands.cs",
    ]
    
    fixed_count = 0
    
    for filepath in interfaces:
        if not os.path.exists(filepath):
            continue
            
        with open(filepath, 'r') as f:
            content = f.read()
        
        original = content
        
        # Pattern: Task<int> MethodName -> int MethodName
        content = re.sub(
            r'Task<int> (\w+)',
            r'int \1',
            content
        )
        
        if content != original:
            with open(filepath, 'w') as f:
                f.write(content)
            print(f"✅ Fixed: {os.path.basename(filepath)}")
            fixed_count += 1
    
    return fixed_count

if __name__ == '__main__':
    print("=" * 60)
    print("CLI Async-to-Sync Refactoring")
    print("=" * 60)
    
    print("\n[1/3] Fixing CommandHelper.cs...")
    fix_command_helper()
    
    print("\n[2/3] Fixing command implementations...")
    num_fixed = fix_cli_command_files()
    print(f"  Total files fixed: {num_fixed}")
    
    print("\n[3/3] Fixing interface signatures...")
    num_fixed = fix_interface_files()
    print(f"  Total interface files fixed: {num_fixed}")
    
    print("\n" + "=" * 60)
    print("✅ Refactoring complete!")
    print("=" * 60)
