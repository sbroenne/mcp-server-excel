---
applyTo: "llm-tests/**"
excludeAgent: "code-review"
---

# LLM evaluations

- Evaluate product discoverability: natural Excel requests, not command/flag
  tutorials. Fix missing guidance in its canonical source, not by coaching the
  test agent through the failing step.
- Use `build_excel_cli_eval`/`build_excel_mcp_eval` and shared budgets from
  `conftest.py`. Normal skill evaluations include the skill; explicitly named
  tool-description-only or tool-availability experiments may omit/restrict it.
- Shared workflows need equivalent CLI/MCP requests and outcome checks.
  Transport assertions and documented entry-point-specific experiments can differ.
- Inspect tool calls and workbook outcomes, not just the final prose. Distinguish
  product failures from auth, Excel, harness, and budget failures. Missing
  prerequisites may skip in the harness; skip/xfail must not hide product defects.
- Setup and commands live in `llm-tests/README.md`. Run affected scenarios only:
  these require Excel and external model access and may incur costs.
