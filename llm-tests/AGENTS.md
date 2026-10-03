# LLM evaluations

Follow the [repository rules](../AGENTS.md). These are implementation and
evaluation instructions; review tasks use the root [Code Review Rules](../AGENTS.md#code-review-rules).

These Python evaluations are on-demand only, outside the normal development
lifecycle. Follow the
[on-demand evaluation policy](../AGENTS.md#build-and-validation); the instructions
below apply only when evaluation work is explicitly requested.

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
- Use the 1.x single-attempt API and native JSON evidence. Do not restore removed
  retries, AI judging/report options, or Azure report prerequisites. `max_turns`
  is advisory; shared time/tool limits are the hard budgets.
- Skill-value comparisons are opt-in, fixed-model, independently scored runs.
  Hold requests, tools, permissions, fixtures, and budgets constant; change only
  the matching skill's availability. Freeze guidance/checkers before paid runs.
  Count every paid smoke/rerun against the approved ceiling.
- Verify successful completed calls and saved/live workbook state. Prove checkers
  reject wrong state before paid runs. Verify actual SDK discovery before paid
  skill comparisons: exactly the matching skill in treatment, none in baseline.
  Call public `pytest_skill_engineering.load_skill` on each actual individual
  directory before SDK probes; SDK-only discovery misses local validation
  failures. Check resolved SDK source paths too, and forbid `send` and
  `send_and_wait`. These probes must make no model requests or start Excel servers.
  Supplied directories are not availability evidence; `empty` mode requires
  explicit `enable_skills=True`. An available but unread skill is not
  demonstrated practical benefit; missing usage is unknown, not zero.
- On-demand setup and commands live in `llm-tests/README.md`. When requested, run
  affected scenarios only; live scenarios require Excel and external model
  access and may incur costs.
