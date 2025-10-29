# Implementation Plan - SuggestedNextActions Refactoring

> **Status**: In Progress  
> **Start Date**: 2025-01-29  
> **Estimated Duration**: 6-10 days  
> **Current Phase**: Phase 0 - Shared Validation Layer

## Overview

This plan implements the design specified in `suggested-next-actions-design.md` to refactor the SuggestedNextActions system with:
- Type-safe action definitions
- Shared validation layer
- Proper layer separation
- Dual format (MCP/CLI) support

## Implementation Phases

### ✅ Phase 0: Shared Validation Layer (Days 1-2)

**Goal**: Create foundation for shared validation across all layers

**Tasks**:
1. ✅ Create `Core/Models/Validation/` directory structure
2. ⏳ Implement `ParameterDefinition.cs` - Parameter validation rules
3. ⏳ Implement `ValidationResult.cs` - Validation result type
4. ⏳ Implement `ActionDefinition.cs` - Action metadata (CLI + MCP)
5. ⏳ Implement `ActionDefinitions.cs` - Central action registry
6. ⏳ Define PowerQuery actions (list, view, import, export, update, delete, refresh, etc.)
7. ⏳ Define Parameter actions (list, get, set, create, delete)
8. ⏳ Define Table actions (list, create, info, rename, delete, etc.)
9. ⏳ Write unit tests for validation logic
10. ⏳ Verify build passes

**Deliverables**:
- [ ] `src/ExcelMcp.Core/Models/Validation/ParameterDefinition.cs`
- [ ] `src/ExcelMcp.Core/Models/Validation/ValidationResult.cs`
- [ ] `src/ExcelMcp.Core/Models/Validation/ActionDefinition.cs`
- [ ] `src/ExcelMcp.Core/Models/Validation/ActionDefinitions.cs`
- [ ] `tests/ExcelMcp.Core.Tests/Unit/Validation/ParameterDefinitionTests.cs`
- [ ] `tests/ExcelMcp.Core.Tests/Unit/Validation/ActionDefinitionTests.cs`

**Success Criteria**:
- All validation tests pass
- Build succeeds with 0 warnings
- ActionDefinitions registry contains all current actions

---

### Phase 1: NextAction Infrastructure (Days 3-4)

**Goal**: Create NextAction abstraction system

**Tasks**:
1. Create `Core/Models/NextActions/` directory
2. Implement `NextActionType.cs` enum
3. Implement `NextActionMcp.cs` - MCP format
4. Implement `NextActionCli.cs` - CLI format
5. Implement `NextAction.cs` base class
6. Implement concrete actions:
   - `ViewItemAction.cs`
   - `ListItemsAction.cs`
   - `CreateItemAction.cs`
   - `UpdateItemAction.cs`
   - `DeleteItemAction.cs`
   - `RefreshItemAction.cs`
   - `ConfigureAction.cs`
   - `DiagnoseAction.cs`
   - `ImportAction.cs`
   - `ExportAction.cs`
7. Implement `NextActionFactory.cs` with domain builders
8. Update `ResultBase` with `NextActions` property
9. Mark `SuggestedNextActions` as `[Obsolete]`
10. Write unit tests for action serialization

**Deliverables**:
- [ ] All NextAction infrastructure files
- [ ] Unit tests with 80%+ coverage
- [ ] Updated ResultBase with backward compatibility

**Success Criteria**:
- All NextAction tests pass
- ToMcp() produces valid JSON structure
- ToCli() produces valid command examples
- Backward compatibility maintained

---

### Phase 2: Core Commands Migration (Days 5-6)

**Goal**: Update Core commands to use NextActionFactory

**Tasks**:
1. Migrate `PowerQueryCommands.cs`
   - Replace string-based suggestions with NextActionFactory calls
   - Context-aware suggestions based on operation result
2. Migrate `ParameterCommands.cs`
3. Migrate `TableCommands/*.cs`
4. Migrate `DataModelCommands/*.cs`
5. Migrate `ScriptCommands.cs` (VBA)
6. Write integration tests
7. Verify all Core tests pass

**Deliverables**:
- [ ] Updated Core command files
- [ ] Integration tests for each command domain
- [ ] All existing tests still pass

**Success Criteria**:
- Core commands populate NextActions instead of SuggestedNextActions
- SuggestedNextActions auto-generates from NextActions
- All Core integration tests pass
- No Core code references CLI command names or MCP action names directly

---

### Phase 3: MCP Server Migration (Days 7-8)

**Goal**: Update MCP tools to serialize NextActions.ToMcp()

**Tasks**:
1. Update `ExcelPowerQueryTool.cs`
   - Use ActionDefinitions for validation
   - Serialize NextActions.ToMcp() instead of string suggestions
2. Update `ExcelParameterTool.cs`
3. Update `TableTool.cs`
4. Update `ExcelDataModelTool.cs`
5. Update `ExcelVbaTool.cs`
6. Update `ExcelConnectionTool.cs`
7. Write MCP integration tests
8. Verify JSON format matches MCP specification

**Deliverables**:
- [ ] Updated MCP tool files
- [ ] MCP integration tests
- [ ] All MCP tests pass

**Success Criteria**:
- MCP tools return structured JSON with tool/action/params
- Validation uses ActionDefinitions (no hardcoded regex)
- All MCP integration tests pass
- JSON format validated against MCP spec

---

### Phase 4: CLI Migration (Days 9-10)

**Goal**: Update CLI commands to display NextActions.ToCli()

**Tasks**:
1. Update CLI `PowerQueryCommands.cs`
   - Use ActionDefinitions for validation
   - Display NextActions.ToCli() with formatted output
2. Update CLI `ParameterCommands.cs`
3. Update CLI `TableCommands.cs`
4. Update CLI `DataModelCommands.cs`
5. Update CLI `ScriptCommands.cs`
6. Write CLI integration tests
7. Manual testing of CLI output

**Deliverables**:
- [ ] Updated CLI command files
- [ ] CLI integration tests
- [ ] Manual test results documented

**Success Criteria**:
- CLI displays full command examples with syntax
- Validation uses ActionDefinitions (no manual arg count checks)
- All CLI integration tests pass
- Output is human-readable and copy-paste ready

---

### Phase 5: Documentation & Cleanup (Day 11)

**Goal**: Update documentation and finalize

**Tasks**:
1. Update README.md with examples
2. Update COMMANDS.md with new format
3. Add migration guide for external consumers
4. Document deprecation timeline for SuggestedNextActions
5. Create v2.0 breaking changes document
6. Final integration testing across all layers
7. Performance testing

**Deliverables**:
- [ ] Updated documentation
- [ ] Migration guide
- [ ] v2.0 breaking changes document
- [ ] Performance benchmarks

**Success Criteria**:
- Documentation complete and accurate
- Migration guide clear for users
- All tests pass (unit + integration)
- Performance acceptable

---

## Progress Tracking

### Completed
- ✅ Design documents (4 comprehensive specs)
- ✅ Validation analysis
- ✅ Layer separation analysis
- ✅ Implementation plan

### In Progress
- ⏳ Phase 0: Shared Validation Layer (Task 1 complete)

### Remaining
- Phase 0: Tasks 2-10
- Phase 1: All tasks
- Phase 2: All tasks
- Phase 3: All tasks
- Phase 4: All tasks
- Phase 5: All tasks

---

## Risk Management

### Known Risks

1. **Backward Compatibility**
   - **Risk**: Breaking existing code that uses SuggestedNextActions
   - **Mitigation**: Deprecated property auto-generates from NextActions
   - **Status**: Mitigated

2. **Performance Impact**
   - **Risk**: NextAction abstraction adds overhead
   - **Mitigation**: Benchmark and optimize if needed
   - **Status**: Monitoring

3. **Testing Coverage**
   - **Risk**: Missing edge cases in validation
   - **Mitigation**: Comprehensive test suite, manual testing
   - **Status**: In progress

4. **Scope Creep**
   - **Risk**: Adding features beyond design
   - **Mitigation**: Strict adherence to design document
   - **Status**: Monitoring

### Mitigation Strategies

- **Incremental commits**: Small, testable changes
- **Continuous testing**: Run tests after each change
- **Documentation**: Keep docs in sync with code
- **Code review**: Request review at phase boundaries

---

## Testing Strategy

### Unit Tests
- Parameter validation logic
- Action definition creation
- NextAction serialization (ToMcp, ToCli)
- Factory methods

### Integration Tests
- Core commands with NextActions
- MCP tools with JSON serialization
- CLI commands with formatted output
- Validation across layers

### Manual Tests
- CLI output readability
- MCP JSON structure
- Performance benchmarks
- Edge cases

---

## Success Metrics

- [ ] 100% of actions defined in ActionDefinitions
- [ ] 80%+ test coverage for new code
- [ ] 0 build warnings
- [ ] All existing tests pass
- [ ] No Core code contains CLI/MCP-specific strings
- [ ] Validation shared across all layers
- [ ] Documentation complete

---

## Notes

- This plan may be adjusted as implementation progresses
- Each phase should be reviewed before proceeding to next
- Breaking changes deferred to v2.0
- Backward compatibility maintained in v1.x

---

## Next Immediate Steps

1. Create `src/ExcelMcp.Core/Models/Validation/` directory
2. Implement `ParameterDefinition.cs`
3. Implement `ValidationResult.cs`
4. Implement `ActionDefinition.cs`
5. Begin defining actions in `ActionDefinitions.cs`
