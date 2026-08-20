---
applyTo: "src/**/*.cs,tests/**/*.cs"
excludeAgent: "code-review"
---

# Bug-fix workflow

1. Trace the failure from the user entry point through generated routing,
   Service, Core, and Excel COM. State the root cause before choosing a fix.
2. Search sibling commands, fallback/retry branches, both entry points, tests,
   guidance, and generated/template sources for the same pattern.
3. Add the smallest regression test that reproduces the bug and observe it fail.
4. Fix the root cause at the owning layer without changing unrelated behavior.
5. Rerun the regression test, then the smallest related test group.
6. Update only documentation, help, descriptions, or workflow hints affected by
   the behavior change.
7. Run the applicable repository audit scripts and summarize the root cause,
   same-pattern findings, behavior change, and validation in the pull request.

There is no fixed quota for tests or documentation files. Coverage must be
proportional to the defect: include the reported case, meaningful boundary or
error cases, and cross-entry-point coverage where the contract crosses CLI/MCP.

Do not:

- Fix only the visible symptom while another branch retains the same defect.
- Assert only `Success`; verify the resulting workbook or response state.
- Add broad catches, silent defaults, or retries without a bounded condition.
- Hand-edit generated artifacts instead of their source.
- Create a separate investigation or summary document in the repository.
