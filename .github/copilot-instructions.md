# GitHub Copilot Instructions - ExcelMcp

> **🎯 Optimized for AI Coding Agents** - Modular, focused instructions for efficient context loading

## 📋 Quick Navigation

**Start here** → Read [CRITICAL-RULES.md](copilot/CRITICAL-RULES.md) first (5 mandatory rules)

**Then reference as needed:**
- 🧪 [Testing Strategy](copilot/testing-strategy.md) - Test architecture, OnDemand pattern, filtering
- 📊 [Excel COM Interop](copilot/excel-com-interop.md) - COM patterns, cleanup, best practices
- 🏗️ [Architecture Patterns](copilot/architecture-patterns.md) - Command pattern, pooling, resource management
- 🧠 [MCP Server Guide](copilot/mcp-server-guide.md) - MCP tools, protocol, error handling
- 🔄 [Development Workflow](copilot/development-workflow.md) - PR process, CI/CD, security, versioning

---

## What is ExcelMcp?

**ExcelMcp** is a Windows-only toolset for programmatic Excel automation via COM interop, designed for coding agents and automation scripts.

**Three Layers:**
1. **Core** (`src/ExcelMcp.Core`) - Excel COM interop business logic
2. **CLI** (`src/ExcelMcp.CLI`) - Command-line interface for scripting
3. **MCP Server** (`src/ExcelMcp.McpServer`) - Model Context Protocol for AI assistants

**Key Capabilities:**
- Power Query M code management (import, export, update, refresh)
- VBA macro management (list, import, export, run)
- Worksheet operations (read, write, create, delete)
- Named range parameters (get, set, create)
- Cell operations (values, formulas)
- Excel instance pooling for MCP server performance

---

## 🎯 Development Quick Start

### Before You Start
1. Read [CRITICAL-RULES.md](copilot/CRITICAL-RULES.md) - 5 mandatory rules
2. Check [Testing Strategy](copilot/testing-strategy.md) for test execution patterns

### Common Tasks
- **Add new command** → Follow patterns in [Architecture Patterns](copilot/architecture-patterns.md)
- **Excel COM work** → Reference [Excel COM Interop](copilot/excel-com-interop.md)
- **Modify pool code** → MUST run OnDemand tests (see [CRITICAL-RULES.md](copilot/CRITICAL-RULES.md))
- **Add MCP tool** → Follow [MCP Server Guide](copilot/mcp-server-guide.md)
- **Create PR** → Follow [Development Workflow](copilot/development-workflow.md)

### Test Execution
```bash
# Development (fast feedback)
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Pre-commit (requires Excel)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"

# Pool cleanup (MANDATORY when modifying pool code)
dotnet test --filter "RunType=OnDemand"
```

---

## 📎 Related Resources

**For Excel automation in other projects:**
- Copy `docs/excel-powerquery-vba-copilot-instructions.md` to your project's `.github/copilot-instructions.md`

**Project Documentation:**
- [Commands Reference](../docs/COMMANDS.md)
- [Architecture Overview](../docs/ARCHITECTURE-REFACTORING.md)
- [Installation Guide](../docs/INSTALLATION.md)
- [Security Improvements](../docs/SECURITY-IMPROVEMENTS.md)

---

## 🔄 Continuous Learning

After completing significant tasks, update these instructions with lessons learned. See [CRITICAL-RULES.md](copilot/CRITICAL-RULES.md) Rule 4.

