---
"excelmcp": patch
---

Calculated-member creation no longer diagnoses generic Excel or provider failures as formula syntax errors. Errors preserve the underlying Excel exception and direct callers to inspect calculated members before retrying; unknown member types are rejected before changing the workbook.

Calculated-member guidance preserves qualified measure names and no longer suggests adding members or sets to Values. PivotChart creation failures retain the created chart, its identity, and the underlying failure instead of deleting it or claiming rollback; PivotChart title reads release the acquired title reference.

OLAP Show Values As readback mismatches now report verification failure rather than claiming Excel ignored the write. Recovery guidance warns that displayed values may already have changed and directs callers to inspect the data; no automatic rollback is performed.
