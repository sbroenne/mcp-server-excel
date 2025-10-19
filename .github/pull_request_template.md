## Summary
Brief description of what this PR does.

## Type of Change
- [ ] 🐛 Bug fix (non-breaking change which fixes an issue)
- [ ] ✨ New feature (non-breaking change which adds functionality)
- [ ] 💥 Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] 📚 Documentation update
- [ ] 🔧 Maintenance (dependency updates, code cleanup, etc.)

## Related Issues
Closes #[issue number]
Relates to #[issue number]

## Changes Made
- Change 1
- Change 2
- Change 3

## Testing Performed
- [ ] Tested manually with various Excel files
- [ ] Verified Excel process cleanup (no excel.exe remains after 5 seconds)
- [ ] Tested error conditions (missing files, invalid arguments, etc.)
- [ ] All existing commands still work
- [ ] VBA script execution tested (if applicable)
- [ ] XLSM file format validation tested (if applicable)
- [ ] VBA trust setup tested (if applicable)
- [ ] Build produces zero warnings

## Test Commands
```bash
# Commands used for testing
ExcelMcp command1 "test.xlsx"
ExcelMcp command2 "test.xlsx" "param"
```

## Screenshots (if applicable)
[Add screenshots showing the new functionality]

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review of code completed
- [ ] Code builds with zero warnings
- [ ] Appropriate error handling added
- [ ] Updated help text (if adding new commands)
- [ ] Updated README.md (if needed)
- [ ] Follows Excel COM best practices from copilot-instructions.md
- [ ] Uses `ExcelHelper.WithExcel()` for Excel operations
- [ ] Properly handles 1-based Excel indexing
- [ ] Escapes user input with `.EscapeMarkup()`
- [ ] Returns consistent exit codes (0 = success, 1+ = error)

## Additional Notes
Any additional information that reviewers should know.