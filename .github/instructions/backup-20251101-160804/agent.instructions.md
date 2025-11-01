# GitHub Coding Agent Instructions (VS Code Agent Mode)

> **Audience:** LLM coding agents (Claude Sonnet 4.5, ChatGPT 5).  
> **Environment:** Visual Studio Code **Agent Mode**.  
> **Goal:** Prefer VS Code built‑in tools and minimize terminal approvals. Support .NET first; also handle npm/Node.js/TypeScript.

---

## 0) Execution Principles (Read First)

- **Always prefer VS Code built‑in tools** over ad‑hoc shell commands.
- **Minimize approval prompts:** Avoid proposing terminal commands that require user confirmation unless there is no viable VS Code equivalent.
- **Respect project context:** Detect whether the workspace is **.NET** or **npm/Node/TypeScript** and apply the corresponding guidance.
- **Be deterministic and quiet:** Use existing tasks, scripts, and launch configs; don’t create new ones unless requested.
- **Short, atomic actions:** When fallback to terminal is unavoidable, batch minimal, high‑value commands and explain why.

---

## 1) Preferences & Fallbacks (At‑a‑Glance)

| Action | **Preferred (VS Code Built‑in)** | **Fallback (Only if necessary)** | Notes |
|---|---|---|---|
| **Run .NET tests** | Testing panel / Test Explorer (run, debug selected, re-run failed) | `dotnet test` | Prefer UI to avoid approvals and get inline diagnostics. |
| **Filter .NET tests by trait** | *(No native trait filter in VS Code Test Explorer yet)* | `dotnet test --filter "Category=Integration"` | Use CLI only when trait filtering is explicitly needed. |
| **Debug .NET app** | VS Code Debugger with `launch.json` | `dotnet run` | Launch configs give breakpoints, variables, no approvals. |
| **Build .NET** | VS Code **Build** task / Command Palette “.NET: Build” | `dotnet build` | Use built‑in build tasks if present. |
| **Run npm scripts** | **NPM Scripts** panel / “Run Script” links | `npm run <script>` | Use scripts explorer to avoid approvals. |
| **Debug Node/TS** | VS Code Debugger with Node launch config | `node …` or `npm start` | Prefer attach/launch configs. |

---

## 2) .NET Workflows
- Use **Testing** view for test runs and debugging.
- Trait filtering fallback:
```bash
dotnet test --filter "Category=Integration"
```

---

## 3) Node.js / TypeScript Workflows
- Use **NPM Scripts** panel and **Debug** view.
- Fallback for tests:
```bash
npm run test -- --runTestsByPath path/to/changed.spec.ts
```

---

## 4) Approval Minimization
- Prefer built‑in UI actions.
- Batch CLI commands if unavoidable.
- Explain necessity before fallback.

---

## 11) Minimal Command Reference
```bash
dotnet test --filter "Category=Integration"
npm run test -- --runTestsByPath path/to/changed.spec.ts
```

---

**End of instructions.**
