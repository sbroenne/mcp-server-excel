---
applyTo: "src/ExcelMcp.Core/Commands/**/*.cs,src/ExcelMcp.McpServer/**/*.cs"
---

# Core Commands Coverage - Mandatory Workflow

> **⚠️ CRITICAL**: When adding Core Commands methods, you MUST expose them in MCP Server

## Quick Reference

| Task | Command | Time |
|------|---------|------|
| Check coverage before commit | `.\scripts\audit-core-coverage.ps1` | 30s |
| Add new Core method | Follow 8-step workflow below | 5-10 min |
| Fix pre-commit hook failure | Add missing enum values + mappings | 2-3 min |
| Verify build | `dotnet build -c Release` | 1-2 min |

---

## Mandatory Workflow: Adding New Core Method

**ALWAYS follow these 8 steps in order:**

```markdown
1. ✅ Add method to Core Commands interface
   File: src/ExcelMcp.Core/Commands/[Feature]/I[Feature]Commands.cs
   Example: Task<OperationResult> NewMethodAsync(IExcelBatch batch);

2. ✅ Implement in Core Commands class  
   File: src/ExcelMcp.Core/Commands/[Feature]/[Feature]Commands.cs

3. ✅ Add enum value to ToolActions.cs
   File: src/ExcelMcp.McpServer/Models/ToolActions.cs
   Example: PowerQueryAction.NewMethod
   ⚠️ Build will show CS8524 error until steps 4-6 complete

4. ✅ Add ToActionString mapping
   File: src/ExcelMcp.McpServer/Models/ActionExtensions.cs
   Example: PowerQueryAction.NewMethod => "new-method",
   ⚠️ CS8524 error persists

5. ✅ Add switch case in MCP Tool
   File: src/ExcelMcp.McpServer/Tools/Excel[Feature]Tool.cs
   Example: PowerQueryAction.NewMethod => await NewMethodAsync(...),
   ⚠️ CS8524 error persists

6. ✅ Implement MCP method
   File: src/ExcelMcp.McpServer/Tools/Excel[Feature]Tool.cs
   Example: private static async Task<string> NewMethodAsync(...)
   ✅ CS8524 errors resolved

7. ✅ Build and verify
   Command: dotnet build -c Release
   Expected: 0 warnings, 0 errors

8. ✅ Update documentation
   Files: MCP prompts, tool descriptions, README (if needed)
```

**Why This Order**: Compiler (CS8524) enforces steps 3-6, preventing you from shipping unexposed Core methods.

---

## Compiler Enforcement (CS8524)

**The compiler FORCES you to expose Core methods** through enum-based switches:

```csharp
// Step 3: Add enum value (compiler checks this)
public enum PowerQueryAction
{
    List,
    View,
    NewMethod  // ⚠️ Forget this → CS8524 error in ActionExtensions.cs
}

// Step 4: Add ToActionString mapping (compiler checks this)
public static string ToActionString(this PowerQueryAction action) => action switch
{
    PowerQueryAction.List => "list",
    PowerQueryAction.View => "view",
    PowerQueryAction.NewMethod => "new-method",  // ⚠️ Forget this → CS8524 error
};

// Step 5: Add switch case in Tool (compiler checks this)
return action switch
{
    PowerQueryAction.List => await ListAsync(...),
    PowerQueryAction.View => await ViewAsync(...),
    PowerQueryAction.NewMethod => await NewMethodAsync(...),  // ⚠️ Forget this → CS8524 error
};
```

**Result**: **Impossible to compile** until all 3 enum mappings are added!

---

## Pre-Commit Hook (Automatic Check)

**Before every commit**, the pre-commit hook runs `audit-core-coverage.ps1` to verify Core methods match enum values.

**Setup** (one-time):
```powershell
.\scripts\pre-commit.ps1
```

**On failure, you see**:
```
❌ Coverage gaps detected! All Core methods must be exposed via MCP Server.

The following interfaces have fewer enum values than Core methods:
  - IRangeCommands: Core has 42 methods, RangeAction has 40 values (missing 2)

Action Required:
  1. Review Core interface for new methods
  2. Add missing enum values to ToolActions.cs
  3. Add ToActionString mappings to ActionExtensions.cs
  4. Add switch cases to appropriate MCP Tools
```

**Fix**: Follow 8-step workflow above.

**Emergency bypass** (use only for non-Core changes):
```bash
git commit --no-verify -m "Message"
```

⚠️ **Never use `--no-verify`** for Core Commands changes - fix the gaps instead!

---

## Manual Coverage Check

**Run anytime** to verify coverage:

```powershell
# Check coverage (shows gaps if any)
.\scripts\audit-core-coverage.ps1

# Check coverage and fail if gaps found (useful in CI/CD)
.\scripts\audit-core-coverage.ps1 -FailOnGaps

# Verbose output with detailed counts
.\scripts\audit-core-coverage.ps1 -Verbose
```

**Expected output when 100% coverage**:
```
Interface           CoreMethods EnumValues Gap Status
---------           ----------- ---------- --- ------
IPowerQueryCommands          18         18   0 ✅
ISheetCommands               13         13   0 ✅
IRangeCommands               42         42   0 ✅
ITableCommands               23         23   0 ✅

Summary: 100% coverage ✅ (156 Core methods, 156 enum values)
```

**When gaps detected**:
```
Interface           CoreMethods EnumValues Gap Status
---------           ----------- ---------- --- ------
IRangeCommands               42         40   2 ❌

Summary: 98.7% coverage (156 Core methods, 154 enum values, 2 gaps)
```

**Fix**: Follow 8-step workflow.

---

## Troubleshooting

### CS8524 Error: "Switch expression does not handle all possible values"

**Cause**: Added enum value but forgot to add it to switch expression.

**Fix**: Add the missing case to the switch expression in the file mentioned in error.

### Pre-Commit Hook Fails with "Coverage gaps detected"

**Cause**: Core interface has more methods than corresponding enum has values.

**Fix**: Follow 8-step workflow (steps 3-6).

### Build Succeeds but Pre-Commit Hook Still Fails

**Cause**: Added Core method but forgot to add enum value.

**Fix**: Add to ToolActions.cs, then mappings in ActionExtensions.cs, then Tool switch case.

---

## Key Takeaways

✅ **Compiler enforces coverage** - CS8524 prevents incomplete implementations  
✅ **Pre-commit hook verifies** - Catches gaps before commit  
✅ **8-step workflow is mandatory** - No shortcuts  
✅ **100% coverage is required** - No exceptions

