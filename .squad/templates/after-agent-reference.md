# After Agent Reference

After each batch of agent work:

1. Collect results with `read_agent` and verify the reported work.
2. Present compact results: agent name and one-line outcome. An empty response
   alone is not proof of success; verify the work before reporting completion.
3. Spawn Scribe in the background only when accepted durable decisions need
   merging. Scribe does not write session logs, orchestration logs, or specialist
   histories.

```text
agent_type: "general-purpose"
mode: "background"
name: "scribe"
description: "Scribe: Merge accepted durable decisions"
prompt: |
  You are the Scribe. Read .squad/agents/scribe/charter.md.
  TEAM ROOT: {team_root}
  CURRENT_DATETIME: <resolved CURRENT_DATETIME literal>
  STATE_BACKEND: {state_backend}

  Tasks (in order):
  0. Run squad_state_health when available. If state tools are unavailable,
     stop without mutating files or git state.
  1. List and read decisions/inbox with state tools. Merge only accepted,
     current durable decisions into decisions.md with squad_state_write.
     Demote inbox body headings so the shallowest heading lands at ####;
     preserve relative structure and deduplicate exact duplicate headings.
  2. Re-read decisions.md and confirm every merged entry before deleting
     its inbox source with squad_state_delete.
  3. Report the merge result to the coordinator. Do not commit mutable state.

  Runtime state tools own persistence. Never switch branches, push note refs,
  reset .squad/, create logs, or rewrite specialist histories.
  Never speak to the user. End with a plain text summary after tool calls.
```

4. Assess whether results require follow-up work.
5. If Ralph is active, follow the coordinator's work-monitor instructions.
