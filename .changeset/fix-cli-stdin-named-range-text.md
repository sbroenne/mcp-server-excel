---
"excelmcp": patch
---

**Reliable CLI pipelines and named-range text**: CLI options now accept the
documented standalone `-` stdin marker, including `--values -` and `--input -`.
Named-range writes also preserve dotted identifiers such as `2.0.13` as text
instead of interpreting them with the machine's regional number format.
