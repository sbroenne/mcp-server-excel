<!--
SYNC IMPACT REPORT - Constitution v2.1.0
========================================

VERSION CHANGE: 2.0.0 → 2.1.0 (MINOR - Material Expansion)
DATE: 2025-11-10
RATIONALE: Expanded Principle XXI from bug-fix-specific documentation to comprehensive documentation synchronization requirement for ALL changes

KEY CHANGES:
- Principle XXI renamed: "Comprehensive Bug Fixing Process" → "Documentation Synchronization"
- Expanded scope: Now applies to ALL changes (features, bug fixes, API changes), not just bug fixes
- Added explicit file list: README.md, McpServer/README.md, server.json, vscode-extension/README.md, gh-pages/index.md
- Clarified rationale: Documentation is first-class artifact, inconsistency causes user confusion
- Updated cross-references in Appendix A

PRINCIPLE XXI - BEFORE (v2.0.0):
- Name: "Comprehensive Bug Fixing Process"
- Scope: Bug fixes only
- Docs requirement: "Update 3+ files: tool/method docs, user docs, LLM prompts"

PRINCIPLE XXI - AFTER (v2.1.0):
- Name: "Documentation Synchronization"  
- Scope: ALL changes (features, bugs, API changes)
- Docs requirement: Minimum 4 files across 6 categories:
  1. Component Documentation (XML docs, [Description] attributes)
  2. User Documentation (5 specific files listed: READMEs, server.json, gh-pages)
  3. LLM Prompts (Prompts/Content/)
  4. Workflow Hints (SuggestedNextActions, error messages)
  5. Verification (consistency check)
  6. PR Description (document updates)

TEMPLATES UPDATED:
✅ Appendix A - Updated cross-reference for Principle XXI
⚠️  bug-fixing-checklist.instructions.md - Should reference new principle name
⚠️  readme-management.instructions.md - Should reference Principle XXI

FILES REQUIRING UPDATES:
⚠️  .github/instructions/bug-fixing-checklist.instructions.md - Update references to "Comprehensive Bug Fixes" → "Documentation Synchronization"
⚠️  .github/instructions/readme-management.instructions.md - Add reference to Principle XXI

FOLLOW-UP TODOS:
- Update bug-fixing-checklist to reference Documentation Synchronization principle
- Update readme-management to reference Principle XXI as authoritative source
- Consider adding documentation sync check to pre-commit hooks

RATIONALE FOR EXPANSION:
Documentation inconsistency is a persistent issue. Users access information from multiple entry points:
- Developers: Main README, NuGet README, source code docs
- LLMs: MCP server.json, prompt files, tool descriptions
- VS Code users: Extension README, marketplace listing
- Web visitors: gh-pages/index.md

All must tell the same story. Scoping documentation requirements to "bug fixes only" was too narrow.
Making this a universal requirement for ALL changes ensures documentation debt never accumulates.
-->

<!--
SYNC IMPACT REPORT - Constitution v2.0.0
========================================

VERSION CHANGE: 1.4.0 → 2.0.0 (MAJOR - Breaking Change)
DATE: 2025-11-10
RATIONALE: Reordered all 25 principles by importance hierarchy to create optimal learning path

KEY CHANGES:
- All 25 principles renumbered based on expert importance assessment
- Principles now organized in 6 tiers: Foundation → Architecture → Quality → Testing → Implementation → Workflow
- Tier headers added throughout principles section
- Preamble updated to explain importance-based ordering
- Appendix A completely updated with new cross-references
- Appendix C reference updated (Principle VIII → Principle X)

TIER STRUCTURE:
- TIER 1 (I-V): Foundation - Catastrophic if violated (Success Flag, COM Management, Test-Before-Commit, NetOffice, MS Docs)
- TIER 2 (VI-VIII): Architecture - System design foundation (Four-Layer, Batch API, Operation Independence)
- TIER 3 (IX-XI): Quality & Process - Discipline enforcement (Integration Testing, PRs, No Placeholders)
- TIER 4 (XII-XVII): Testing Discipline - Test quality and isolation
- TIER 5 (XVIII-XXI): Implementation Standards - Code quality (MCP JSON, Tool Descriptions, Enums, Bug Fixes)
- TIER 6 (XXII-XXV): Workflow & Policy - Process optimization

BREAKING CHANGE: All principle numbers changed
- III. Test-Before-Commit (was VI)
- IV. NetOffice Reference First (was XXV)
- V. Modern .NET/MS Docs (was XXIV)
- VI. Four-Layer Architecture (was III)
- VII. Batch API (was IV)
- VIII. Operation Independence (was V)
- IX. Integration-First Testing (was VII)
- X. Pull Request Discipline (was VIII)
- XI. No Placeholders (was IX)
- XVIII. MCP JSON (was XII)
- XXI. Comprehensive Bug Fixes (was XI)

TEMPLATES UPDATED:
✅ Appendix A - All 25 principle cross-references updated
✅ Appendix C - Principle VIII → Principle X reference updated
✅ .github/copilot-instructions.md - "25 principles enforced" reference verified
✅ All instruction files validated - No specific principle number references found

FILES REQUIRING NO UPDATES:
✅ .github/instructions/*.instructions.md - No principle number references found
✅ Source code (*.cs) - No principle number references found
✅ specs/**/*.md - No principle number references found
✅ docs/**/*.md - No principle number references found

FOLLOW-UP TODOS:
- None required - All cross-references and templates already updated

RATIONALE FOR REORDERING:
New developers need foundational concepts first (Success Flag, COM Management, Test-Before-Commit).
Expert priority assessment: catastrophic-if-violated principles must be understood before supporting details.
Testing minutiae (naming standards) can come later as implementation details.
Creates optimal learning path: Foundation → Implementation → Supporting Details
-->

# ExcelMcp Constitution

**Version**: 2.1.0 | **Ratified**: 2025-11-10

## Preamble

This Constitution establishes the foundational principles and governance framework for ExcelMcp, a Windows-only toolset for programmatic Excel automation via COM interop designed for coding agents and automation scripts.

These principles are ordered by importance from FOUNDATION → ARCHITECTURE → QUALITY → TESTING → IMPLEMENTATION → WORKFLOW. All principles are NON-NEGOTIABLE and apply to all contributions, modifications, and deployments.

---

## Core Principles

### Tier 1: Foundation (Catastrophic if Violated)

### I. Success Flag Integrity (CRITICAL)

**Description**: NEVER set `Success = true` when `ErrorMessage` is set. Invariant: `Success == true` ⟹ `ErrorMessage == null || ErrorMessage == ""`.

**Rationale**: LLMs and automation tools rely on the Success flag to determine operation outcomes. Violated invariant causes workflow failures, silent data corruption, and wasted debugging time. This is the most fundamental contract in the system - if this fails, everything built on top fails silently.

**Implementation Requirements**:
- Set `Success = true` ONLY in try block after actual success
- Set `Success = false` ALWAYS in catch block  
- Pre-commit hook `check-success-flag.ps1` MUST pass
- Code review MUST verify every `Success = ` assignment

**Cross-Reference**: Rule 1 (critical-rules), PowerQuerySuccessErrorRegressionTests

---

### II. COM Resource Management

**Description**: All COM objects MUST be released explicitly in finally blocks. Zero COM leaks enforced by pre-commit hook.

**Rationale**: COM leaks cause Excel processes to hang indefinitely, consume memory without releasing, corrupt workbooks through stale references, and create file lock issues preventing cleanup. This is the foundation of stability for COM interop - without proper resource management, the entire system becomes unreliable. Explicit release in finally blocks guarantees cleanup even when exceptions occur.

**Implementation Requirements**:
- Declare COM objects as `dynamic? obj = null`
- Use try-finally blocks with `ComUtilities.Release(ref obj!)` in finally
- Pre-commit hook `check-com-leaks.ps1` MUST report 0 leaks
- Exception: Session management files (ExcelSession.cs, ExcelBatch.cs) - reviewed separately
- All intermediate COM objects released (ranges, collections, sheets, queries)

**Cross-Reference**: Rule 5 (critical-rules), excel-com-interop, ComUtilities.cs

---

### III. Test-Before-Commit (NON-NEGOTIABLE)

**Description**: NEVER commit, push, or create PRs without first running tests for the code you changed. Build MUST pass with zero warnings.

**Rationale**: This is the primary quality gate. Prevents breaking changes from reaching main branch, eliminates waste from debugging preventable failures, enforces CI/CD discipline, and respects team time. Without this, all other quality principles become meaningless.

**Implementation Requirements**:
- Build MUST pass with zero warnings
- Feature-specific tests MUST pass before commit
- Pre-commit hooks MUST pass
- Session/batch code changes require OnDemand tests

**Cross-Reference**: Rule 0 (critical-rules), CONTRIBUTING.md, DEVELOPMENT.md

---

### IV. NetOffice Reference First

**Description**: MUST check NetOffice GitHub repository (https://github.com/NetOfficeFw/NetOffice) for working sample code BEFORE implementing new Excel COM Interop features AND while troubleshooting COM issues.

**Rationale**: NetOffice has a long, successful history of working with Excel in .NET and provides strongly-typed C# wrappers for ALL Office COM APIs. The repository contains proven patterns for every Excel COM operation (ranges, worksheets, workbooks, charts, PivotTables, Power Query, VBA, connections, formatting, etc.). Studying working implementations from NetOffice prevents common COM pitfalls (1-based indexing, object cleanup, async issues, variant types, OLE automation errors). Real-world examples from this mature project are more reliable than documentation alone and save significant debugging time.

**Implementation Requirements**:
- ALWAYS search NetOffice repository before implementing new Excel COM automation features
- ALWAYS search NetOffice when troubleshooting COM issues or unexpected behavior
- Use NetOffice search for specific COM objects/methods: "PivotTable CreatePivotTable", "QueryTable Refresh", "Range.Value2"
- Study NetOffice patterns for dynamic interop conversion and proper COM object handling
- Document NetOffice references in code comments for complex COM patterns
- Prefer NetOffice proven patterns over untested approaches when both available
- NEVER skip NetOffice search - it is the most comprehensive Excel COM reference available

**Cross-Reference**: Rule 9 (critical-rules), excel-com-interop, OlapPivotTableFieldStrategy.cs (NetOffice reference example)

---

### V. Modern .NET and Microsoft Documentation First

**Description**: MUST use modern .NET 8 best practices and C# language features. MUST check Microsoft documentation via Microsoft Docs MCP Server before implementing .NET, COM Interop, or Azure features.

**Rationale**: Microsoft documentation is more up-to-date and authoritative than general LLM training data. LLM knowledge cutoffs can be months or years old, missing critical updates, security patches, breaking changes, and new best practices. The Microsoft Docs MCP Server provides real-time access to official documentation, ensuring implementations use current patterns, avoid deprecated APIs, and follow recommended practices. This is especially critical for COM Interop and Azure services which evolve rapidly.

**Implementation Requirements**:
- ALWAYS query Microsoft Docs MCP Server (microsoft_docs_search) before implementing .NET, COM Interop, or Azure features
- Use modern C# language features (.NET 8 compatible): records, pattern matching, async/await, nullable reference types
- Follow official .NET API guidelines and framework design guidelines
- COM Interop: Verify patterns against Microsoft Office Interop documentation
- Azure services: Check latest SDK versions and migration guides
- Document Microsoft Docs references in code comments for complex patterns
- Prefer official Microsoft examples over third-party tutorials when conflicts arise

**Cross-Reference**: excel-com-interop, architecture-patterns, Microsoft Docs MCP Server, .NET 8 documentation

---

### Tier 2: Architecture (System Design)

### VI. Four-Layer Architecture

**Description**: Maintain strict separation between layers. Dependencies MUST flow downward only: CLI/MCP Server → Core → ComInterop.

**Rationale**: Layered architecture is the foundational design pattern that enables everything else. It prevents circular dependencies, enables component reuse across CLI and MCP Server, facilitates isolated testing, maintains clear separation of concerns, and allows independent deployment/versioning of components. Without this architecture, the system becomes a tangled mess. Production code referencing test code violates architectural boundaries.

**Implementation Requirements**:
- ComInterop layer NEVER references higher layers
- Core references ComInterop only
- CLI/MCP Server reference Core + ComInterop
- Production code NEVER references test projects
- No InternalsVisibleTo from production to tests

**Cross-Reference**: Rule 11 (critical-rules), architecture-patterns

---

### VII. Batch API Architecture

**Description**: ALL Excel operations MUST use IExcelBatch for exclusive workbook access and session management.

**Rationale**: Batch API is the core architectural pattern for all Excel operations. It provides 75-90% performance improvement, guaranteed COM cleanup, exclusive file access preventing corruption, and consistent error handling. This is the operational foundation that makes everything else work efficiently and reliably.

**Implementation Requirements**:
- All Core Command methods accept `IExcelBatch batch` as first parameter
- Explicit `await batch.SaveAsync()` required for persistence
- Tests use batch API (no ValueTask.FromResult wrapper)
- Batch ID optional for grouping operations in MCP Server

**Cross-Reference**: Rule 3 (critical-rules), architecture-patterns, ExcelSession.cs, ExcelBatch.cs

---

### VIII. Operation Independence and Separation of Concerns

**Description**: Core operations MUST NOT depend on other Core operations to perform their function. Shared logic MUST be refactored to base/helper classes or utilities, not implemented via cross-operation dependencies.

**Rationale**: 
Cross-operation dependencies severely complicate test isolation and violate fundamental software engineering principles. This is critical for maintainability and testability. Operations must be independently testable units that follow Single Responsibility Principle and Dependency Inversion Principle.

By extracting shared logic to helper classes/utilities:
- **Test isolation**: Each operation can be tested independently without complex setup or mocking
- **Modularity**: Helper classes can be tested separately, establishing a clear test hierarchy
- **Maintainability**: Changes to shared logic don't require modifying multiple operation classes
- **Dependency Inversion**: Operations depend on stable abstractions (helper utilities), not on other concrete operations
- **Clear boundaries**: Helper classes define explicit contracts for shared functionality

**Implementation Requirements**:
- Core Command constructors MUST have NO dependencies on other Core Command interfaces
- Shared functionality MUST be extracted to helper classes, utilities, or private methods
- Test hierarchy: Helper classes/utilities FIRST (unit tests), Core Commands SECOND (integration tests)
- NEVER inject `IXxxCommands` interfaces into Command constructors
- Extract common logic to helper class H when operations A and B need same functionality

**Cross-Reference**: Rule 11 (critical-rules), architecture-patterns, testing-strategy, specs/QUERYTABLE

---

### Tier 3: Quality & Process (Discipline Enforcement)

### IX. Integration-First Testing Philosophy

**Description**: No traditional unit tests. Integration tests against real Excel ARE our unit tests.

**Rationale**: Excel COM API cannot be meaningfully mocked. Integration tests provide real-world validation with executable documentation. This is an accepted Architecture Decision (ADR-001). This principle defines the entire testing philosophy and must be understood early.

**Implementation Requirements**:
- Each test creates unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`
- NEVER share test files between tests
- Binary assertions only (NO "accept both" patterns)
- ALWAYS verify actual Excel state after operations (round-trip validation)
- SaveAsync ONLY for persistence tests
- Session/batch modifications require OnDemand tests

**Cross-Reference**: ADR-001-NO-UNIT-TESTS.md, Rule 12 (critical-rules), testing-strategy

---

### X. Pull Request Discipline

**Description**: ALL changes via pull requests. NEVER commit directly to main branch.

**Rationale**: Pull requests are the enforcement mechanism for all other principles. Direct commits bypass CI/CD pipelines, skip code review, violate branch protection, and eliminate quality gates. This principle enables all quality assurance.

**Implementation Requirements**:
- Create feature branch before making changes
- Build MUST pass with zero warnings
- All relevant tests MUST pass
- Documentation MUST be updated
- Automated review comments MUST be fixed BEFORE human review
- Pre-commit hook blocks direct commits to main

**Cross-Reference**: Rule 6 (critical-rules), Rule 19 (PR review automation), development-workflow, CONTRIBUTING.md

---

### XI. No Placeholders or TODOs

**Description**: Code MUST be complete before commit. NO TODO, FIXME, HACK, or XXX markers. NO NotImplementedException.

**Rationale**: This principle enforces done-is-done philosophy. Placeholders accumulate technical debt invisibly, confuse contributors about work status, block pre-commit hooks, create false sense of progress, and leave incomplete features in production. Every commit must represent complete, production-ready functionality. Incomplete work stays in branches until finished.

**Implementation Requirements**:
- All features fully implemented with real Excel COM operations
- Passing tests for all implemented functionality
- Complete documentation (tool docs, user docs, prompts)
- No commented-out code (use git history)
- Pre-commit hook blocks commits with TODO/FIXME/HACK/XXX markers
- NotImplementedException NEVER allowed

**Cross-Reference**: Rule 2, Rule 8 (critical-rules), PRE-COMMIT-SETUP.md

---

### Tier 4: Testing Discipline (Test Quality)

### XII. Round-Trip Validation

**Description**: Tests MUST verify actual Excel state after create/update operations, not just success flag. For operations that replace content, verify content was replaced AND old content is gone.

**Rationale**: "Operation completed" ≠ "Operation did the right thing". This is fundamental to test effectiveness. Bug reports showed UpdateAsync was merging M code instead of replacing it. Tests passed because they only checked Success=true, not actual content. Round-trip validation proves operations do what they claim.

**Implementation Requirements**:
- After CREATE: Verify object exists via List/Get
- After UPDATE: Verify changes applied AND old content gone
- After DELETE: Verify object removed
- For REPLACE operations: Verify new content present AND old content absent
- Test multiple sequential updates to expose merging bugs

**Cross-Reference**: testing-strategy, bug reports

---

### XIII. Test File Isolation

**Description**: Each test MUST create unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`. NEVER share test files between tests.

**Rationale**: Shared test files cause file lock issues, race conditions in parallel execution, unpredictable test order dependencies, cascading failures from corrupted shared state. File isolation ensures tests are independent, reproducible, and parallelizable. This is essential for reliable test execution.

**Implementation Requirements**:
- Use `CoreTestHelper.CreateUniqueTestFileAsync(className, methodName, tempDir, extension)`
- Use `IClassFixture<TempDirectoryFixture>` for temp directory management
- Extension ".xlsx" default, ".xlsm" for VBA tests
- Unique file pattern: `{className}_{methodName}_{timestamp}.xlsx`

**Cross-Reference**: Rule 12 (critical-rules), testing-strategy, CoreTestHelper

---

### XIV. Binary Assertions

**Description**: Tests MUST use binary assertions (True/False). NEVER use "accept both" patterns or conditional assertions.

**Rationale**: "Accept both" assertions hide real failures, provide false confidence, make debugging harder, violate test determinism. Tests should ALWAYS have single expected outcome. If multiple outcomes valid, test is poorly designed or testing wrong thing. Deterministic tests are essential for maintainability.

**Implementation Requirements**:
- Use `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- Use `Assert.False(result.Success)` for expected failure tests
- NEVER use patterns like `if (success || failure)` or `Assert.True(a || b)`
- ALWAYS verify actual Excel state after operations
- Provide failure message explaining what was expected

**Cross-Reference**: Rule 12 (critical-rules), testing-strategy

---

### XV. Single Batch Testing

**Description**: Tests MUST use only ONE batch per test and MUST NOT call SaveAsync unless explicitly testing persistence. We know Excel can save files - no need to test Microsoft's functionality.

**Rationale**: SaveAsync is slow (2-5 seconds), forces unnecessary disk I/O, disrupts Excel COM state when called mid-test. 90% of tests verify in-memory state only. Removing unnecessary saves makes test suite 50%+ faster. Test efficiency matters for development velocity. Persistence testing reserved for explicit round-trip scenarios.

**Implementation Requirements**:
- Only ONE batch per test (`await using var batch`)
- NO SaveAsync in middle of test (breaks COM state and subsequent operations)
- NO SaveAsync for tests verifying in-memory state only
- SaveAsync ONLY for test fixtures or explicit persistence tests (create → save → reopen → verify)
- Performance: 2-5 seconds eliminated per removed SaveAsync

**Cross-Reference**: Rule 14 (critical-rules), testing-strategy

---

### XVI. Test Scope Discipline

**Description**: ALWAYS run tests ONLY for the specific code you modified. NEVER run all integration tests during development.

**Rationale**: Integration tests require Excel COM automation and are SLOW (10+ minutes for full suite). Running all tests wastes time and resources during development. Feature-specific tests provide fast feedback (3-5 minutes). Fast feedback loops improve developer productivity. Full test suite runs in CI/CD pipeline only.

**Implementation Requirements**:
- Use Feature trait to target specific test groups
- Run OnDemand tests ONLY when modifying session/batch code
- Document which tests were run in commit message
- Full integration suite runs in CI/CD only

**Cross-Reference**: Rule 16 (critical-rules), testing-strategy, DEVELOPMENT.md

---

### XVII. Test Naming Standard

**Description**: Test names MUST follow pattern: MethodName_StateUnderTest_ExpectedBehavior. NO "Async" suffix, NO generic states.

**Rationale**: Consistent naming makes tests discoverable, self-documenting, easier to understand failures. Pattern identifies: what's tested (method), conditions (state), outcome (behavior). Generic names provide no diagnostic value when tests fail. Clear naming is documentation that stays current.

**Implementation Requirements**:
- MethodName: Command being tested (no "Async" suffix)
- StateUnderTest: Specific scenario (NOT "Valid", "InvalidFile", "Success")
- ExpectedBehavior: Clear outcome (Returns*, Creates*, Removes*, Throws*)
- Use descriptive states: EmptyWorkbook, DuplicateName, NonActiveSheet

**Cross-Reference**: Rule 12 (critical-rules), testing-strategy, TEST-NAMING-STANDARD.md

---

### Tier 5: Implementation Standards (Code Quality)

### XVIII. MCP JSON Response Pattern

**Description**: MCP tools MUST return JSON for ALL responses. NEVER throw exceptions for business logic errors. Throw McpException ONLY for protocol errors.

**Rationale**: MCP specification defines the contract with AI agents. Violations break agent workflows. MCP requires business errors return JSON with `isError: true` flag. HTTP 200 + JSON error enables clients to parse and handle gracefully. Throwing exceptions for business errors violates MCP spec and makes agents fail unpredictably. Core Commands return result objects with Success flag - serialize them directly.

**Implementation Requirements**:
- ALWAYS return `JsonSerializer.Serialize(result, JsonOptions)` for Core Command results
- Let `result.Success` flag indicate business errors
- Throw `McpException` only for parameter validation, file not found, batch not found
- NEVER throw `McpException` for table not found, query failed, connection error
- Validate parameters early before calling Core Commands

**Cross-Reference**: Rule 17 (critical-rules), mcp-server-guide

---

### XIX. Tool Description Accuracy

**Description**: Tool [Description] attributes are part of MCP schema sent to LLMs. They MUST be accurate, current, and document server-specific behavior.

**Rationale**: LLMs use tool descriptions for server-specific guidance and workflow planning. Inaccurate descriptions cause wrong defaults, incorrect workflows, and user confusion. Descriptions are ALWAYS visible when LLMs browse tools. This is the interface contract with AI agents.

**Implementation Requirements**:
- Verify when changing tools: purpose clear, server behavior documented, performance guidance accurate
- Explain non-enum parameter values (loadDestination, formatCode, etc.)
- Don't duplicate schema info (types, required flags auto-provided)
- Update BOTH tool description AND prompt file when changing behavior

**Cross-Reference**: Rule 18 (critical-rules), mcp-server-guide, mcp-llm-guidance

---

### XX. Complete Enum Mappings

**Description**: Every enum value MUST have a mapping in ToActionString(). Missing mappings cause unhandled ArgumentException at runtime.

**Rationale**: Missing enum mappings cause MCP Server to throw exceptions instead of returning JSON, violating MCP protocol. Runtime exceptions harder to debug than compile-time errors. Complete mappings ensure predictable behavior and proper error handling.

**Implementation Requirements**:
- When adding enum value, add ToActionString() mapping immediately
- Use switch expression with exhaustive cases
- Throw ArgumentException in default case
- Code review MUST verify completeness
- Regression tests for all enum mappings

**Cross-Reference**: Rule 15 (critical-rules), ToActionString extensions

---

### XXI. Documentation Synchronization

**Description**: ALL changes (features, bug fixes, API changes) MUST synchronize documentation across ALL user-facing artifacts before PR approval. Minimum 4 files: component docs, user docs, LLM prompts, and deployment artifacts.

**Rationale**: Inconsistent documentation causes user confusion, incorrect LLM workflows, and wasted support time. Documentation is a first-class artifact, not an afterthought. Users consume information from multiple entry points (READMEs, MCP schema, VS Code extension, web docs) - all must tell the same story. Outdated documentation is worse than no documentation.

**Implementation Requirements**:
1. **Component Documentation** - Update XML docs, [Description] attributes, inline comments at implementation site
2. **User Documentation** - Update ALL relevant files:
   - `/README.md` - Main project documentation
   - `/src/ExcelMcp.McpServer/README.md` - NuGet package docs
   - `/src/ExcelMcp.McpServer/server.json` - MCP server manifest (tools, prompts)
   - `/vscode-extension/README.md` - VS Code extension docs
   - `/gh-pages/index.md` - Website landing page
   - Component-specific docs in relevant directories
3. **LLM Prompts** - Update prompt files in `src/ExcelMcp.McpServer/Prompts/Content/` for tool/workflow changes
4. **Workflow Hints** - Update SuggestedNextActions, error messages, WorkflowHint in operation results
5. **Verification** - All documentation files list matching capabilities, no contradictions, examples work
6. **PR Description** - Document which files updated and why

**Cross-Reference**: Rule 13 (critical-rules), readme-management, mcp-llm-guidance, bug-fixing-checklist

---

### Tier 6: Workflow & Policy (Process Optimization)

### XXII. PR Review Automation

**Description**: After creating PR, ALWAYS check automated review comments from Copilot and GitHub Advanced Security within 1-2 minutes. Fix ALL automated issues before requesting human review.

**Rationale**: Automated reviewers catch code quality and security issues early. Fixing promptly improves quality, reduces human reviewer workload, speeds approval, prevents technical debt. Human reviewers should focus on architecture and design, not style issues that machines can catch.

**Implementation Requirements**:
- Create PR then immediately retrieve comments via `gh api` or mcp_github tool
- Fix ALL automated issues in single commit
- Request human review ONLY after automated issues resolved
- Common issues: improper inheritdoc, inefficiency patterns, nullable access, nested ifs

**Cross-Reference**: Rule 19 (critical-rules), development-workflow, PR #139

---

### XXIII. COM API Transparency

**Description**: ExcelMcp MUST NOT add security layers or policy enforcement on top of Excel's COM API. Security is the caller's responsibility.

**Rationale**: ExcelMcp is a thin automation wrapper over Excel's COM API, not a security framework. Adding security policies (like enforcing `SavePassword = false`) creates false security expectations, breaks transparency with underlying COM API, and prevents legitimate use cases where callers need control over COM properties. Security decisions belong at the application/caller layer, not in the automation library. This maintains clear responsibility boundaries.

**Implementation Requirements**:
- NEVER set COM properties to enforce security policies
- MCP tools MAY expose COM API parameters but MUST NOT enforce default security values
- Documentation MUST clarify security is caller's responsibility
- Connection strings, passwords, sensitive data are pass-through from COM API
- Sanitization/masking MUST be at display/logging layer, not Core commands

**Cross-Reference**: excel-connection-types-guide, architecture-patterns, specs/001-connection-management

---

### XXIV. No Automatic Commits

**Description**: NEVER commit or push code automatically. All commits, pushes, and merges MUST require explicit user approval.

**Rationale**: Automatic commits prevent user review, bypass pre-commit validation, violate audit requirements, and introduce unintended changes. Explicit approval ensures users maintain control and understand what's being committed. This is fundamental to user trust and control.

**Implementation Requirements**:
- All automated tools/scripts/agents MUST prompt for user approval before commit/push/merge
- No background or silent commits allowed
- Display changes clearly before requesting approval
- Document this rule in all agent instructions

**Cross-Reference**: Rule 21 (critical-rules), agent instructions

---

### XXV. VS Code Tools First

**Description**: Coding agents MUST use VS Code tools (replace_string_in_file, read_file, grep_search, file_search, etc.) before resorting to shell commands (run_in_terminal with PowerShell/cmd).

**Rationale**: VS Code tools provide more efficient workflow, significantly reduce approval prompts for users, and increase development velocity. This is especially critical for Claude, which prefers shell commands by default. Shell commands require user approval for each execution, creating friction and slowing development. VS Code tools operate within the editor context, providing better integration and requiring fewer approvals.

**Implementation Requirements**:
- File operations MUST use read_file, replace_string_in_file, create_file instead of Get-Content, Set-Content, New-Item
- Code searches MUST use grep_search or semantic_search instead of Select-String or findstr
- File discovery MUST use file_search instead of Get-ChildItem or dir
- Error checking MUST use get_errors instead of parsing dotnet build output
- Shell commands reserved for: build operations (dotnet build/test), git operations, pre-commit hooks
- Document rationale when shell command is necessary despite VS Code tool availability

**Cross-Reference**: Agent instructions, VS Code extension documentation, MCP tool usage guidelines

---

## Appendices

### Appendix A: Principle Cross-References to Critical Rules

| Principle | Critical Rules | Source Documents |
|-----------|----------------|------------------|
| I. Success Flag Integrity | Rule 1 | critical-rules, PowerQuerySuccessErrorRegressionTests |
| II. COM Resource Management | Rule 5 | critical-rules, excel-com-interop |
| III. Test-Before-Commit | Rule 0 | critical-rules, CONTRIBUTING.md, DEVELOPMENT.md |
| IV. NetOffice Reference First | Rule 9 | critical-rules, excel-com-interop, NetOffice repository |
| V. Modern .NET and Microsoft Docs First | N/A | excel-com-interop, architecture-patterns, Microsoft Docs MCP, .NET 8 docs |
| VI. Four-Layer Architecture | Rule 11 | architecture-patterns, critical-rules |
| VII. Batch API Architecture | Rule 3 | critical-rules, architecture-patterns, ExcelSession.cs |
| VIII. Operation Independence | Rule 11 | critical-rules, architecture-patterns, testing-strategy, specs/QUERYTABLE |
| IX. Integration-First Testing | Rule 12 | ADR-001, testing-strategy, critical-rules |
| X. Pull Request Discipline | Rule 6, Rule 19 | critical-rules, development-workflow, CONTRIBUTING.md |
| XI. No Placeholders or TODOs | Rule 2, Rule 8 | critical-rules, PRE-COMMIT-SETUP.md |
| XII. Round-Trip Validation | Rule 12 | testing-strategy, bug reports |
| XIII. Test File Isolation | Rule 12 | critical-rules, testing-strategy, CoreTestHelper |
| XIV. Binary Assertions | Rule 12 | testing-strategy, critical-rules |
| XV. Single Batch Testing | Rule 14 | testing-strategy, critical-rules, user requirement 2025-11-10 |
| XVI. Test Scope Discipline | Rule 16 | critical-rules, testing-strategy, DEVELOPMENT.md |
| XVII. Test Naming Standard | Rule 12 | TEST-NAMING-STANDARD.md, testing-strategy |
| XVIII. MCP JSON Response Pattern | Rule 17 | critical-rules, mcp-server-guide |
| XIX. Tool Description Accuracy | Rule 18 | critical-rules, mcp-server-guide, mcp-llm-guidance |
| XX. Complete Enum Mappings | Rule 15 | critical-rules, RangeAction enum |
| XXI. Documentation Synchronization | Rule 13 | readme-management, mcp-llm-guidance, bug-fixing-checklist, critical-rules |
| XXII. PR Review Automation | Rule 19 | critical-rules, development-workflow, PR #139 |
| XXIII. COM API Transparency | N/A | excel-connection-types-guide, architecture-patterns, specs/001 |
| XXIV. No Automatic Commits | Rule 21 | critical-rules, agent instructions |
| XXV. VS Code Tools First | N/A | Agent instructions, VS Code documentation, MCP tool guidelines |

### Appendix C: Governance Process

**Amendment Authority**:
- Constitution maintained by project maintainers
- Community feedback via GitHub Issues/Discussions
- Changes require PR approval per Principle X (Pull Request Discipline)

**Version Numbering**:
- **MAJOR**: Backward incompatible principle removals or redefinitions
- **MINOR**: New principle added or materially expanded guidance  
- **PATCH**: Clarifications, wording, typo fixes

**Enforcement**:
- Pre-commit hooks enforce Rules 0, 1, 5, 8, 21
- Code review enforces all principles
- CI/CD pipeline validates compliance
- Violations require documented justification and maintainer approval

**Sync Process**:
- Constitution changes documented in Sync Impact Report (HTML comment)
- Templates and documentation updated to reflect principle changes
- Cross-references maintained in Appendix A
