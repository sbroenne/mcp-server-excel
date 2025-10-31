# Testing Architecture: Before & After Azure Self-Hosted Runners

## Current State (GitHub-Hosted Runners Only)

```
┌─────────────────────────────────────────────────────────────────────┐
│ GitHub Repository: sbroenne/mcp-server-excel                        │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ GitHub Actions: build-mcp-server.yml, build-cli.yml                 │
│ Runner: windows-latest (GitHub-hosted)                              │
│ Has: Windows Server, .NET 8, PowerShell                             │
│ MISSING: Microsoft Excel ❌                                          │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Test Execution                                                       │
│                                                                      │
│  ✅ UNIT TESTS (46 tests)                                           │
│     dotnet test --filter "Category=Unit"                            │
│     - ComUtilitiesTests                                             │
│     - ExcelBatchTests                                               │
│     - ExcelSessionTests                                             │
│     - McpServer routing tests                                       │
│     Result: ✅ PASS (~2-5 seconds)                                  │
│                                                                      │
│  ❌ INTEGRATION TESTS (91 tests) - SKIPPED                          │
│     Reason: Require Microsoft Excel                                 │
│     - RangeCommandsTests (values, formulas, editing)                │
│     - PowerQueryCommandsTests                                       │
│     - VbaCommandsTests                                              │
│     - ConnectionCommandsTests                                       │
│     - DataModelCommandsTests                                        │
│     Result: ❌ NOT RUN                                              │
│                                                                      │
│  Total Automated Coverage: 46/137 tests (34%) ⚠️                    │
└─────────────────────────────────────────────────────────────────────┘
```

## Future State (GitHub-Hosted + Azure Self-Hosted)

```
┌─────────────────────────────────────────────────────────────────────┐
│ GitHub Repository: sbroenne/mcp-server-excel                        │
└─────────────────────────────────────────────────────────────────────┘
           │                                    │
           │ Unit Tests                         │ Integration Tests
           ▼                                    ▼
┌──────────────────────────────┐  ┌──────────────────────────────────┐
│ GitHub-Hosted Runner         │  │ Azure Self-Hosted Runner         │
│                              │  │                                  │
│ Windows-latest               │  │ Azure Windows VM                 │
│ Has:                         │  │ Has:                             │
│  - Windows Server            │  │  - Windows Server 2022           │
│  - .NET 8 SDK                │  │  - .NET 8 SDK                    │
│  - PowerShell                │  │  - PowerShell                    │
│  - Build tools               │  │  - Microsoft Excel ✅            │
│                              │  │  - Office 365 License            │
│ Cost: FREE                   │  │ Cost: ~$30-65/month              │
└──────────────────────────────┘  └──────────────────────────────────┘
           │                                    │
           ▼                                    ▼
┌──────────────────────────────┐  ┌──────────────────────────────────┐
│ UNIT TESTS                   │  │ INTEGRATION TESTS                │
│                              │  │                                  │
│ Category=Unit                │  │ Category=Integration             │
│ 46 tests                     │  │ 91 tests                         │
│                              │  │                                  │
│ ✅ ExcelMcp.ComInterop       │  │ ✅ RangeCommands                 │
│ ✅ ExcelMcp.Core (unit)      │  │   - Get/Set Values               │
│ ✅ ExcelMcp.CLI (unit)       │  │   - Formulas                     │
│ ✅ ExcelMcp.McpServer (unit) │  │   - Editing (insert/delete)      │
│                              │  │   - Search (find/replace)        │
│ Runtime: 2-5 seconds         │  │   - Discovery (UsedRange)        │
│ Frequency: Every PR/push     │  │   - Hyperlinks                   │
│                              │  │                                  │
│                              │  │ ✅ PowerQueryCommands            │
│                              │  │   - Import/Export M code         │
│                              │  │   - Create/Update/Delete queries │
│                              │  │   - Refresh operations           │
│                              │  │                                  │
│                              │  │ ✅ VbaCommands                   │
│                              │  │   - Import/Export VBA modules    │
│                              │  │   - Run macros                   │
│                              │  │                                  │
│                              │  │ ✅ ConnectionCommands            │
│                              │  │   - OLEDB/ODBC/Text/Web          │
│                              │  │   - Connection properties        │
│                              │  │                                  │
│                              │  │ ✅ DataModelCommands             │
│                              │  │   - DAX measures                 │
│                              │  │   - Table relationships          │
│                              │  │                                  │
│                              │  │ Runtime: 13-15 minutes           │
│                              │  │ Frequency: Nightly + Manual      │
└──────────────────────────────┘  └──────────────────────────────────┘
           │                                    │
           └────────────┬───────────────────────┘
                        ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Total Automated Coverage: 137/137 tests (100%) ✅                   │
│                                                                      │
│ Coverage Improvement: +91 tests (+197% increase)                    │
└─────────────────────────────────────────────────────────────────────┘
```

## Workflow Execution Flow

```
┌─────────────────────────────────────────────────────────────────────┐
│ Trigger Event                                                        │
│  - Scheduled (cron: '0 2 * * *') - Nightly at 2 AM UTC             │
│  - Manual (workflow_dispatch) - On-demand via Actions tab          │
│  - Optional PR (commented out) - Can enable to block merges         │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 1: Checkout code                                               │
│   uses: actions/checkout@v4                                         │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 2: Setup .NET                                                  │
│   uses: actions/setup-dotnet@v4 (dotnet-version: 8.0.x)            │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 3: Display Excel Version (verification)                        │
│   PowerShell: New-Object -ComObject Excel.Application              │
│   Output: "✅ Excel Version: 16.0" (example)                        │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 4: Restore dependencies                                        │
│   dotnet restore                                                    │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 5: Build                                                       │
│   dotnet build --no-restore --configuration Release                │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 6-9: Run Integration Tests (4 projects)                        │
│                                                                      │
│  6. ExcelMcp.Core.Tests                                             │
│     Filter: Category=Integration&RunType!=OnDemand                  │
│     Expected: ~60 tests                                             │
│                                                                      │
│  7. ExcelMcp.CLI.Tests                                              │
│     Filter: Category=Integration&RunType!=OnDemand                  │
│     Expected: ~10 tests                                             │
│                                                                      │
│  8. ExcelMcp.McpServer.Tests                                        │
│     Filter: Category=Integration&RunType!=OnDemand                  │
│     Expected: ~10 tests                                             │
│                                                                      │
│  9. ExcelMcp.ComInterop.Tests                                       │
│     Filter: Category=Integration&RunType!=OnDemand                  │
│     Expected: ~11 tests                                             │
│                                                                      │
│ Total Runtime: ~13-15 minutes                                       │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 10: Upload Test Results (if: always())                         │
│   Artifact: integration-test-results-{run_number}                   │
│   Path: **/TestResults/*.trx                                        │
│   Retention: 30 days                                                │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Step 11-12: Cleanup (if: always())                                  │
│                                                                      │
│  11. Kill Excel processes                                           │
│      Get-Process excel | Stop-Process -Force                        │
│                                                                      │
│  12. Check for orphaned processes (warning only)                    │
│      Detect any remaining Excel.exe processes                       │
│                                                                      │
│ Purpose: Prevent Excel process leaks on runner                      │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────────┐
│ Workflow Complete                                                    │
│   - Test results available in Artifacts                             │
│   - Excel processes cleaned up                                      │
│   - Runner ready for next run                                       │
└─────────────────────────────────────────────────────────────────────┘
```

## Key Differences: Unit vs Integration Tests

| Aspect | Unit Tests | Integration Tests |
|--------|-----------|-------------------|
| **Runner** | GitHub-hosted (windows-latest) | Azure self-hosted |
| **Excel Required** | ❌ No | ✅ Yes |
| **Test Count** | 46 tests | 91 tests |
| **Runtime** | 2-5 seconds | 13-15 minutes |
| **Test Scope** | Logic, parsing, routing | Excel COM operations |
| **Frequency** | Every PR/push | Nightly + manual |
| **Cost** | FREE | ~$30-65/month |
| **Examples** | ComUtilities, ExcelBatch | Range operations, Power Query |
| **Failures** | Block PRs | Alert only (initially) |

## Benefits of Hybrid Approach

1. **Fast Feedback** - Unit tests run in seconds on every PR
2. **Comprehensive Coverage** - Integration tests validate real Excel behavior
3. **Cost Effective** - Only pay for Azure runner when needed
4. **Flexible** - Can schedule integration tests or run manually
5. **Non-Blocking** - Integration tests don't block PRs (initially)
6. **Scalable** - Can add more runners if needed

## Testing Strategy Summary

```
Development Workflow:
├─ Developer makes changes
├─ Push to feature branch
├─ PR created → Unit tests run (GitHub-hosted) → Fast feedback (2-5 sec)
├─ PR approved and merged to main
├─ Nightly: Integration tests run (Azure self-hosted) → Full coverage (15 min)
└─ If integration tests fail → Alert team → Fix in next PR
```

**Result:** Best of both worlds - fast unit tests + comprehensive integration tests!
