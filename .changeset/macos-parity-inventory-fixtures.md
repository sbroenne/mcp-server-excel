---
"excelmcp": minor
---

Generate an action-level macOS capability inventory from Core contracts and use
it for production routing. Classify Power Query, VBA, and unsupported scenario
operations as explicit macOS limitations because Apple Events and Office.js do
not expose APIs that satisfy their public contracts. Remove the signed VBA
helper architecture, candidate gates, packaging, and trust workflow.

Refine the generated execution plan with Excel 16.113.1 dictionary evidence:
workbook connection actions and QueryTable creation are now limitation
candidates rather than misleading native candidates because the installed
dictionary exposes neither the required workbook connection collection nor a
QueryTable construction command.
