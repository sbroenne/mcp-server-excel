---
workflow: product-launch-video
flow: automation
storyboard: no
message: "Understand what Excel MCP does, how it differs, and whether you need it."
destination: youtube-and-github-pages
aspect: 1920x1080
language: en
length: 120s
audience: people who have never heard of Excel MCP Server
angle: one synthetic sales workbook, from raw data to decisions
narration: yes
voice: af_heart
music: none
---

## Intent

A polished two-minute introduction replacing the website's existing intro.
Adapt the sister PowerPoint MCP video's white editorial canvas, black ink,
Segoe UI / Cascadia Mono typography and purposeful transitions to Excel green.
Warm female English narration, no music. The story and autonomous build were
confirmed before setup. Final preview approval is still required before render.

The opening must answer a first-time viewer's questions: what is Excel MCP
Server, why use real Excel, how is it different from file-based tools, and when
is it the right choice? Explain the trade-off honestly: file-based tools can be
better for simple exports and environments without Windows and desktop Excel.

## Assets

- Sister project: sbroenne/mcp-server-powerpoint, videos/powerpoint-mcp-intro — visual and implementation reference.
- Repository screenshots: excel-demo-table-chart.png and excel-demo-python.png — genuine existing product evidence.
- A new synthetic sales workbook — safe, reproducible Excel evidence.

## Customizations

- Power Query, Data Model/DAX, PivotTables, charts, VBA, Python in Excel,
  real calculation, and preservation of workbook features.
- MCP and excelcli receive equal treatment, especially in the closing frame.
- Verbatim, quiet overlay captions. Genuine screenshots plus clearly labeled
  explanatory code; never present invented UI or unexecuted code as captured output.

## Notes

- No website crawl: the reference aesthetic and source material are supplied.
- Do not change Excel security settings or touch unrelated user workbooks.
- No music, no generic gradients, no idle animation, no external telemetry reports.
- The user approved rendering for YouTube on September 10, 2026. Include the
  qualified CLI token-efficiency benefit. Deliver an MP4 and captions for upload;
  replace the website embed after the new YouTube URL is available.
- Comparison source, reviewed September 10, 2026:
  https://github.com/anthropics/skills/blob/fa0fa64bdc967915dc8399e803be67759e1e62b8/skills/xlsx/SKILL.md.
  The published skill uses openpyxl for editing and LibreOffice for calculation;
  do not imply that it cannot calculate formulas at all.
- Fact review is recorded in `evidence.json`. Do not claim that macro execution
  preserves all security settings: it changes application AutomationSecurity.
  Project editing still requires manually enabled VBA trust. Python runs in
  Microsoft's cloud and requires supported Microsoft 365 and internet access.
