# ADR-001: Test at the boundary that owns the behavior

**Status:** Current

## Context and decision

ExcelMcp contains both Excel-dependent behavior and ordinary .NET logic.
Workbook operations are tested against desktop Excel, normally through the
shared Service. Parsing, serialization, mapping, and generation have focused
Excel-independent tests. Adapter tests cover what CLI and MCP add to that
shared behavior.

## Reasons and tradeoffs

Mocking Excel is faster and easier to run anywhere, but cannot establish COM
marshaling, refresh completion, application state, or workbook persistence.
Requiring Excel for a pure parser would add process cost without testing a
more realistic boundary. Repeating the entire workbook matrix through both
adapters would add cost without isolating their responsibilities.

This split keeps fast feedback for isolated logic while retaining evidence
from the real application. It also means an Excel-free hosted CI run cannot
establish that workbook automation works.

## Implementation and guidance

- [Excel-independent parsing tests](../tests/ExcelMcp.Core.Tests/Unit/ServiceRegistryJsonParsingTests.cs)
- [Service workbook tests](../tests/ExcelMcp.Service.Tests/PersistentServicePowerQueryReadContractTests.cs)
- [Test-writing instructions](../tests/AGENTS.md) and [commands and fixtures](../tests/README.md)
