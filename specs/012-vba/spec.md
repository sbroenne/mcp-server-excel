# Feature Specification: VBA Macro Management

**Feature Branch**: `012-vba`  
**Created**: 2024-01-10  
**Status**: ✅ **IMPLEMENTED**  
**Last Updated**: 2025-11-10

## Implementation Status

**✅ FULLY IMPLEMENTED** - All VBA macro operations functional.

**✅ Implemented:**
- ✅ List VBA modules
- ✅ View VBA code
- ✅ Import VBA from .bas files
- ✅ Export VBA to .bas files
- ✅ Delete VBA modules
- ✅ Run VBA procedures (with parameters)
- ✅ Update VBA code

**Code Location:** `src/ExcelMcp.Core/Commands/Vba/`

## User Scenarios

### User Story 1 - Version Control VBA Code (Priority: P1) 🎯 MVP

As a developer, I need to export VBA modules to files for version control.

**Acceptance Scenarios**:
1. **Given** module "Module1" with VBA code, **When** I export, **Then** .bas file created
2. **Given** 5 modules, **When** I export all, **Then** 5 .bas files created
3. **Given** exported VBA, **When** I diff in git, **Then** changes visible

### User Story 2 - Import and Update VBA (Priority: P1) 🎯 MVP

As a developer, I need to import VBA from files and update existing modules.

**Acceptance Scenarios**:
1. **Given** .bas file, **When** I import, **Then** module appears in workbook
2. **Given** existing module, **When** I update from file, **Then** code replaced
3. **Given** updated VBA, **When** I save, **Then** changes persist

### User Story 3 - Run VBA Procedures (Priority: P2)

As a developer, I need to execute VBA macros programmatically with parameters.

**Acceptance Scenarios**:
1. **Given** procedure "ProcessData", **When** I run with params, **Then** macro executes
2. **Given** long-running macro, **When** I run with timeout, **Then** timeout enforced
3. **Given** macro that fails, **When** I run, **Then** error message returned

### User Story 4 - List and View VBA (Priority: P2)

As a developer, I need to discover what VBA modules exist and view their code.

**Acceptance Scenarios**:
1. **Given** workbook with 3 modules, **When** I list, **Then** I see all 3 with line counts
2. **Given** module name, **When** I view, **Then** I see full VBA code

## Requirements

### Functional Requirements
- **FR-001**: List all VBA modules with names and line counts
- **FR-002**: View VBA code from any module
- **FR-003**: Import VBA from .bas files
- **FR-004**: Export VBA to .bas files
- **FR-005**: Delete VBA modules
- **FR-006**: Run VBA procedures with optional parameters
- **FR-007**: Update existing VBA module code
- **FR-008**: Timeout support for long-running macros (default 30 seconds)

### Non-Functional Requirements
- **NFR-001**: VBA operations require .xlsm file format
- **NFR-002**: VBA Trust settings must allow programmatic access
- **NFR-003**: Run operations timeout after configurable duration
- **NFR-004**: COM object cleanup guaranteed

## Success Criteria
- ✅ All 7 VBA methods implemented
- ✅ Integration tests cover all operations (requires VBA trust)
- ✅ Version control workflows functional

## Technical Context

### Excel COM API
- `Workbook.VBProject` - VBA project object
- `VBProject.VBComponents` - Module collection
- `VBComponent.CodeModule` - Code access
- `VBComponent.Export()` - Save to .bas file
- `VBComponents.Import()` - Load from .bas file
- `Application.Run()` - Execute procedure

### Architecture
- VbaCommands with partial classes (Lifecycle, Operations)
- Batch API for exclusive access
- Timeout support for Run operations
- VBA Trust detection

### Key Design Decisions

#### Decision 1: Require .xlsm Format
**Rationale**: VBA macros only persist in macro-enabled workbooks
**Implementation**: Tests use .xlsm extension, validation checks file type

#### Decision 2: VBA Trust Required
**Detection**: Attempt to access VBProject, catch COM exception if trust disabled
**Why**: Excel security blocks programmatic VBA access by default

#### Decision 3: Run with Timeout
**Default**: 30 seconds
**Implementation**: `RunAsync(batch, "Module.Procedure", timeout, params)`
**Why**: Prevents infinite loops from hanging automation

## Testing Strategy
- **Tests**: `VbaCommandsTests.cs`
- **Coverage**: List, View, Import, Export, Delete, Run, Update
- **Requirements**: VBA Trust must be enabled (manual setup or policy)
- **Edge Cases**: Modules with syntax errors, missing procedures, timeout scenarios

### VBA Trust Setup
Tests require: "Trust access to the VBA project object model" enabled in Excel Trust Center

## Related Documentation
- **Testing**: `testing-strategy.instructions.md`
- **Excel COM**: `excel-com-interop.instructions.md`
- **Security**: VBA trust configuration in docs
