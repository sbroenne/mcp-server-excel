#!/usr/bin/env python3
"""
Remove remaining .MethodAsync( calls - convert to .Method(
"""
import re
from pathlib import Path

def fix_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original = content
    
    # Remove await keywords before _coreCommands calls
    content = re.sub(r'await\s+_coreCommands\.', r'_coreCommands.', content)
    
    # Remove Async suffix from method names
    content = re.sub(r'_coreCommands\.(\w+)Async\(', r'_coreCommands.\1(', content)
    
    # Remove async method declarations that don't have await
    content = re.sub(r'private\s+async\s+Task\s+(\w+)', r'private void \1', content)
    content = re.sub(r'public\s+async\s+Task\s+(\w+)', r'public void \1', content)
    
    # Remove async method declarations with generics
    content = re.sub(r'private\s+async\s+Task<(\w+)>\s+(\w+)', r'private \1 \2', content)
    content = re.sub(r'public\s+async\s+Task<(\w+)>\s+(\w+)', r'public \1 \2', content)
    
    if content != original:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

def main():
    cli_commands_dir = Path('src/ExcelMcp.CLI/Commands')
    
    files = list(cli_commands_dir.glob('*.cs'))
    files = [f for f in files if 'Interface' not in f.name and 'CommandHelper' not in f.name]
    
    fixed_count = 0
    for filepath in files:
        if fix_file(filepath):
            print(f"Fixed: {filepath.name}")
            fixed_count += 1
    
    print(f"\nTotal fixed: {fixed_count} files")

if __name__ == '__main__':
    main()
