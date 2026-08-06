---
"excelmcp": patch
---

Stop transient MSBuild workers after pre-commit release validation and exclude release-only skill generation from the benchmark dependency graph so subsequent builds do not inherit or create locked build-task assemblies.
