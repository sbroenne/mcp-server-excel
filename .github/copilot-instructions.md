# GitHub Copilot Instructions - ExcelMcp

> **🎯 Optimized for AI Coding Agents** - Modular, path-specific instructions following GitHub Copilot best practices

## 📋 Quick Navigation

**Start here** → Read [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) first (5 mandatory rules)

**Path-Specific Instructions** (auto-applied based on file context):
- 🧪 [Testing Strategy](instructions/testing-strategy.instructions.md) - Test architecture, OnDemand pattern, filtering
- 📊 [Excel COM Interop](instructions/excel-com-interop.instructions.md) - COM patterns, cleanup, best practices
- 🏗️ [Architecture Patterns](instructions/architecture-patterns.instructions.md) - Command pattern, pooling, resource management
- 🧠 [MCP Server Guide](instructions/mcp-server-guide.instructions.md) - MCP tools, protocol, error handling
- 🔄 [Development Workflow](instructions/development-workflow.instructions.md) - PR process, CI/CD, security, versioning

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
1. Read [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) - 5 mandatory rules
2. Check [Testing Strategy](instructions/testing-strategy.instructions.md) for test execution patterns

### Common Tasks
- **Add new command** → Follow patterns in [Architecture Patterns](instructions/architecture-patterns.instructions.md)
- **Excel COM work** → Reference [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- **Modify pool code** → MUST run OnDemand tests (see [CRITICAL-RULES.md](instructions/critical-rules.instructions.md))
- **Add MCP tool** → Follow [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- **Create PR** → Follow [Development Workflow](instructions/development-workflow.instructions.md)

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

After completing significant tasks, update these instructions with lessons learned. See [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) Rule 4.

---

## 📚 How Path-Specific Instructions Work

GitHub Copilot automatically loads instructions based on the files you're working with:

- Working in `tests/**/*.cs`? → [Testing Strategy](instructions/testing-strategy.instructions.md) auto-applies
- Working in `src/ExcelMcp.Core/**/*.cs`? → [Excel COM Interop](instructions/excel-com-interop.instructions.md) auto-applies
- Working in `src/ExcelMcp.McpServer/**/*.cs`? → [MCP Server Guide](instructions/mcp-server-guide.instructions.md) auto-applies
- Working in `.github/workflows/**/*.yml`? → [Development Workflow](instructions/development-workflow.instructions.md) auto-applies
- **All files** → [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) always applies

This modular approach ensures you get relevant context without overwhelming the AI with unnecessary information.

