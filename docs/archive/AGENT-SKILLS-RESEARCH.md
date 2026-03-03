# Agent Skills Research for AI Coding Assistants

> **Comprehensive reference for implementing skills across VS Code/GitHub Copilot, Claude Code, and other AI coding assistants**

## Table of Contents

1. [Overview](#overview)
2. [VS Code / GitHub Copilot Skills](#vs-code--github-copilot-skills)
3. [Claude Code Skills](#claude-code-skills)
4. [Cross-Platform Skills (npx skills)](#cross-platform-skills-npx-skills)
5. [Other AI Assistants](#other-ai-assistants)
6. [Comparison Matrix](#comparison-matrix)
7. [Best Practices](#best-practices)

---

## Overview

Agent Skills are reusable instruction sets that extend AI coding assistants with domain-specific knowledge, workflows, and capabilities. They enable consistent, reliable behavior when working with specific tools, frameworks, or codebases.

### Key Concepts

| Term | Definition |
|------|------------|
| **Skill** | A self-contained instruction set with metadata (SKILL.md) |
| **Agent** | The AI coding assistant (Copilot, Claude Code, Cursor, etc.) |
| **Instructions** | Context-specific guidance files (.instructions.md) |
| **Prompts** | Reusable prompt templates (.prompt.md) |
| **Commands** | Slash commands for specific workflows |

---

## VS Code / GitHub Copilot Skills

### Directory Structure

```
~/.copilot/skills/
├── {skill-name}/
│   ├── SKILL.md              # Main skill definition (required)
│   └── references/           # Optional supporting files
│       ├── api-reference.md
│       └── examples.md
```

### SKILL.md Format

```yaml
---
name: your-skill-name
description: Brief description of what this Skill does and when to use it
license: MIT
version: 1.0.0
tags:
  - keyword1
  - keyword2
repository: https://github.com/owner/repo
documentation: https://docs.example.com
---

# Your Skill Name

## Instructions
Provide clear, step-by-step guidance for the agent.

## Examples
Show concrete examples of using this Skill.

## Tool Map
List related tools and when to use them.

## Reference Documentation
- references/api-reference.md
- references/examples.md
```

### Frontmatter Fields

| Field | Required | Description |
|-------|----------|-------------|
| `name` | Yes | Unique identifier (kebab-case recommended) |
| `description` | Yes | Brief description shown to users |
| `license` | No | License identifier (e.g., MIT, Apache-2.0) |
| `version` | No | Semantic version (e.g., 1.0.0) |
| `tags` | No | Array of keywords for discovery |
| `repository` | No | Source code repository URL |
| `documentation` | No | Documentation website URL |

### VS Code Settings

```json
{
  "chat.useAgentSkills": true,
  "github.copilot.chat.codeGeneration.useInstructionFiles": true
}
```

### Instructions Files (.instructions.md)

Located in `.github/instructions/` or `.vscode/` directories:

```markdown
---
applyTo: "**/*.ts,**/*.tsx"
---
# TypeScript Coding Standards

## Guidelines
- Use TypeScript for all new code
- Prefer interfaces over type aliases
- Use strict null checks
```

#### applyTo Patterns

| Pattern | Applies To |
|---------|------------|
| `"**"` | All files |
| `"**/*.ts"` | TypeScript files |
| `"src/**/*.py"` | Python files in src/ |
| `"docs/**/*.md"` | Markdown in docs/ |

### Prompt Files (.prompt.md)

Reusable prompts with metadata:

```markdown
---
agent: 'agent'
model: Claude Sonnet 4
tools: ['githubRepo', 'search/codebase']
description: 'Generate a new React form component'
---
Your goal is to generate a new React form component...
```

### Custom Agents

Located in `.github/agents/` or `.copilot/agents/`:

```yaml
---
name: Planner
displayName: Implementation Planner
description: Generate an implementation plan for features
tools: ['fetch', 'githubRepo', 'search', 'usages']
model: Claude Sonnet 4
handoffs:
  - label: Implement Plan
    agent: agent
    prompt: Implement the plan outlined above.
    send: false
---

# Planning Instructions
You are in planning mode. Generate an implementation plan...
```

---

## Claude Code Skills

### Directory Structure

```
.claude/skills/
├── {skill-name}/
│   ├── SKILL.md              # Main skill definition (required)
│   ├── REFERENCE.md          # Optional reference docs
│   └── scripts/              # Optional utility scripts
│       └── validate.py

# Or global skills:
~/.claude/skills/
└── {skill-name}/
    └── SKILL.md
```

### SKILL.md Format

```yaml
---
name: your-skill-name
description: Brief description with trigger terms. Use when working with X or Y.
allowed-tools: Read, Grep, Glob, Bash(python:*)
user-invocable: true
context: fork
hooks:
  PreToolUse:
    - matcher: "Bash"
      hooks:
        - type: command
          command: "./scripts/security-check.sh"
---

# Your Skill Name

## Instructions
1. First step
2. Second step

## Examples
Show concrete usage examples.
```

### Frontmatter Fields

| Field | Required | Description |
|-------|----------|-------------|
| `name` | Yes | Unique skill identifier |
| `description` | Yes | Description including trigger terms |
| `allowed-tools` | No | Restrict available tools (comma-separated or array) |
| `user-invocable` | No | `true` (default) = appears in slash menu, `false` = model-only |
| `context` | No | `fork` = isolated sub-agent context |
| `hooks` | No | Hook configurations for the skill lifecycle |

### Tool Restrictions

```yaml
# Comma-separated
allowed-tools: Read, Grep, Glob

# Array format
allowed-tools:
  - Read
  - Grep
  - Glob
  - Bash(python:*)  # Only allow python commands
```

### File Imports

Reference other files within SKILL.md:

```markdown
See @README for project overview.
For details, see @docs/api-reference.md.

# Import from home directory
- @~/.claude/my-project-instructions.md
```

### Slash Commands

Located in `.claude/commands/`:

```markdown
---
description: Review code for quality issues
allowed-tools: Bash(git add:*), Bash(git status:*)
hooks:
  PreToolUse:
    - matcher: "Bash"
      hooks:
        - type: command
          command: "./scripts/pre-review.sh"
---

## Context
- Current git status: !`git status`
- Current branch: !`git branch --show-current`

## Your task
Review the staged changes for code quality issues.
```

### CLAUDE.md (Project Instructions)

Root-level project configuration:

```markdown
# Project Instructions

## Overview
Brief project description.

## Development Setup
- Required tools and versions
- Environment setup

## Coding Standards
- Style guidelines
- Naming conventions

## Testing
- How to run tests
- Coverage requirements
```

### Sub-Agents

Located in `.claude/agents/`:

```yaml
---
name: code-reviewer
description: Review code for quality and best practices
tools: Read, Grep, Glob
model: sonnet
permissionMode: default
skills: pr-review, security-check
---

You are a code reviewer. Analyze code for:
1. Code organization
2. Error handling
3. Security concerns
4. Test coverage
```

---

## Cross-Platform Skills (npx skills)

### Installation

```powershell
# From GitHub shorthand
npx skills add vercel-labs/agent-skills

# From specific skill
npx skills add vercel-labs/agent-skills --skill frontend-design

# Install globally
npx skills add vercel-labs/agent-skills --global

# Install for specific agents
npx skills add vercel-labs/agent-skills -a claude-code -a cursor
```

### Supported Agents

| Agent | Project Directory | Global Directory |
|-------|-------------------|------------------|
| `claude-code` | `.claude/skills/` | `~/.claude/skills/` |
| `cursor` | `.cursor/skills/` | `~/.cursor/skills/` |
| `github-copilot` | `.copilot/skills/` | `~/.copilot/skills/` |
| `codex` | `.codex/skills/` | `~/.codex/skills/` |
| `gemini-cli` | `.gemini/skills/` | `~/.gemini/skills/` |
| `opencode` | `.opencode/skills/` | `~/.opencode/skills/` |
| `windsurf` | `.windsurf/skills/` | `~/.windsurf/skills/` |
| `kilo` | `.kilo/skills/` | `~/.kilo/skills/` |
| `goose` | `.goose/skills/` | `~/.goose/skills/` |
| **And 34+ more** | See `npx skills add --help` | See `npx skills add --help` |

### Skill Discovery Locations

```
# Search priority (in order):
1. Root directory (if contains SKILL.md)
2. skills/, skills/.curated/, skills/.experimental/
3. .claude/skills/, .cursor/skills/, .opencode/skills/
4. Recursive search (fallback)
```

---

## Other AI Assistants

### Cursor

**Configuration file:** `.cursorrules` (root of project)

```markdown
# Project Rules
- Prefer using yarn
- Generated commit messages should be in English
- Use TypeScript strict mode
```

**Additional locations:**
- `.cursor/skills/` - Skills directory
- `.cursor/instructions/` - Instructions files

### Windsurf/Codeium

**Configuration:** `.windsurf/skills/` directory

Uses similar SKILL.md format to other agents.

### Gemini CLI

**Configuration:** `.gemini/skills/` directory

Follows the npx skills specification format.

---

## Comparison Matrix

| Feature | GitHub Copilot | Claude Code | Cursor | Windsurf |
|---------|---------------|-------------|--------|----------|
| **Skills Directory** | `~/.copilot/skills/` | `.claude/skills/` | `.cursor/skills/` | `.windsurf/skills/` |
| **Main File** | `SKILL.md` | `SKILL.md` | `SKILL.md` | `SKILL.md` |
| **Instructions** | `.instructions.md` | CLAUDE.md + imports | `.cursorrules` | `.instructions.md` |
| **Commands** | `.prompt.md` | `.claude/commands/` | N/A | N/A |
| **Custom Agents** | `.agent.yaml` | `.claude/agents/` | N/A | N/A |
| **Tool Restrictions** | Via settings | `allowed-tools` | N/A | N/A |
| **Hooks** | N/A | Full lifecycle | N/A | N/A |
| **MCP Support** | Yes | Yes | Limited | Limited |
| **applyTo Patterns** | Yes | Via description | N/A | N/A |

---

## Best Practices

> **Source:** Official documentation from [agentskills.io](https://agentskills.io/specification), [Claude Code](https://code.claude.com/docs/en/skills), [VS Code Copilot](https://code.visualstudio.com/docs/copilot/copilot-customization), and [vercel-labs/skills](https://github.com/vercel-labs/skills) - researched 2026-02-02.

### Official Specification Requirements

From **agentskills.io/specification**:

| Field | Required | Constraints |
|-------|----------|-------------|
| `name` | Yes | 1-64 chars, lowercase alphanumeric + hyphens, no leading/trailing hyphens, no `--` |
| `description` | Yes | 1-1024 chars, describes WHAT + WHEN to use |
| `license` | No | License name or reference to bundled file |
| `compatibility` | No | 1-500 chars, environment requirements (OS, packages, network) |
| `metadata` | No | Key-value map for custom properties |
| `allowed-tools` | No | Space-delimited pre-approved tools (experimental) |

### SKILL.md Size Guidelines (OFFICIAL)

From **agentskills.io** and **Claude Code docs**:

| Layer | Token Budget | Content |
|-------|-------------|---------|
| Metadata | ~100 tokens | `name` + `description` (loaded at startup for ALL skills) |
| Instructions | **< 5000 tokens** | Full SKILL.md body (loaded when skill activated) |
| Resources | As needed | Files in `scripts/`, `references/`, `assets/` |

**Official recommendation:** Keep SKILL.md **under 500 lines**. Move detailed reference material to separate files.

### Progressive Disclosure Pattern

From **agentskills.io**:

1. **Discovery:** Agents load only `name` + `description` at startup
2. **Activation:** When task matches description, agent reads full SKILL.md
3. **Execution:** Agent follows instructions, loading referenced files as needed

**Why this matters:** Keep main skill lean. Don't front-load everything.

### Effective Descriptions (Official Examples)

From **agentskills.io/specification**:

```yaml
# Good - describes WHAT + WHEN with keywords
description: Extracts text and tables from PDF files, fills PDF forms, and merges multiple PDFs. Use when working with PDF documents or when the user mentions PDFs, forms, or document extraction.

# Bad - vague
description: Helps with PDFs.
```

### Using the `compatibility` Field

From **agentskills.io** - use this for platform/environment requirements:

```yaml
---
name: excel-automation
description: Automate Excel workbooks with Power Query, VBA, and data operations.
compatibility: Requires Windows + Microsoft Excel 2016+ (COM interop). Does NOT work on macOS/Linux.
---
```

### Directory Structure (Official)

From **agentskills.io**:

```
skill-name/
├── SKILL.md           # Required - instructions + metadata
├── scripts/           # Optional - executable code
├── references/        # Optional - detailed documentation
└── assets/            # Optional - templates, resources
```

### Skill Content Types (Claude Code)

From **code.claude.com/docs/en/skills**:

| Type | Purpose | Invocation |
|------|---------|------------|
| **Reference** | Knowledge Claude applies to current work (conventions, patterns) | Auto-loaded when relevant |
| **Task** | Step-by-step instructions for specific actions (deploy, commit) | Manual via `/skill-name` |

Use `disable-model-invocation: true` for tasks you want manual control over.

### Cross-Platform Compatibility

Skills work across 38+ agents including:
- GitHub Copilot: `.github/skills/` or `~/.copilot/skills/`
- Claude Code: `.claude/skills/` or `~/.claude/skills/`
- Cursor: `.cursor/skills/`
- OpenCode, Codex, Windsurf, Roo Code, etc.

**To maximize compatibility:**
1. Use only standard frontmatter fields (`name`, `description`, `compatibility`)
2. Keep instructions in standard Markdown
3. Avoid agent-specific features in shared content
4. Test with `npx skills add <repo> --list` before publishing

### CLI Skills: Emphasize Discovery

For CLI-based skills, don't document every parameter. Emphasize `--help`:

```markdown
## Core Principle
Use `--help` for parameters. `excelcli <command> --help` is the authoritative source.
```

**Why:** CLI help is always current; documentation gets stale.

---

## References

- [VS Code Copilot Customization](https://code.visualstudio.com/docs/copilot/customization)
- [Claude Code Skills Documentation](https://code.claude.com/docs/en/skills)
- [skills CLI](https://github.com/vercel-labs/skills) (formerly add-skill)
- [MCP Server Protocol](https://modelcontextprotocol.io/)

---

## Changelog

| Date | Change |
|------|--------|
| 2026-02 | Updated `add-skill` references to `skills` (command renamed), expanded agent support list |
| 2026-02 | Added SKILL.md size guidelines, requirements visibility, CLI discovery patterns |
| 2025-01 | Initial research document |
