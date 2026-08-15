---
"excelmcp": patch
---

**More reliable CLI automation** (#764, #765, #766, #767): the CLI now rejects unknown options, reports all missing required parameters together, keeps visible sessions genuinely visible, and cleans up Excel when a timed-out daemon is forced to stop. Python in Excel polling now respects the session timeout, while the generated CLI skill ships complete live command help and shared domain guidance.
