# Mistakes to avoid

Use [working safely with Excel](behavioral-rules.md) for the shared rules.
Table, Data Model, Power Query, PivotTable, and slicer examples below require
Windows. The experimental Mac beta has a
[gated subset](https://excelmcpserver.dev/macos-support/); do not retry an
unsupported or uncertain mutation through another backend.

| Mistake | Better approach |
|---------|-----------------|
| Rewriting a whole range to change one cell | Write that cell; preserve formulas elsewhere |
| Deleting and rebuilding a Table to update data | Append, resize, or update the existing object |
| Formatting Table headers as plain cells | Use the Table's style |
| Formatting PivotTable cells with range styling | Use field number formats; leave unsupported visual styling alone |
| Reapplying each style property separately | Send all properties together |
| Repeated discovery without a changed state | Reuse known names; inspect again after relevant changes or errors |
| Asking again for permission already given | Proceed within scope; ask only about unresolved choices |
| Inventing a file path or session ID | Use the user's path and a returned matching session |
| Treating a failed call as undo | Inspect partial effects before retrying |
| Saving at the end of a continue-on-error batch | Stop on failure and save only after all intended work succeeds |
| Leaving a failed batch session open | Explicitly close a session owned by that job without saving |
| Using a worksheet table as if it were in the Data Model | Add/load it to the model first |
| Editing source rows without refreshing downstream data | Refresh the model when needed, then the PivotTable |
| Multiplying PivotTable totals for line-item revenue | Calculate each source row, or use DAX `SUMX` |
| Assuming Table slicers filter an independent PivotTable | Filter each intended object explicitly |

## Recovering a failed Power Query creation

Creation adds the query before loading its data. A load error can leave that
query and some load objects behind. Repeating `create` then fails with
"already exists"; deleting everything risks unrelated data and dependencies.

1. Inspect the query list, the surviving query's load configuration, and its
   destination.
2. Evaluate corrected M code where useful, and check its actual preview.
3. If the query exists, update it. Create only if the query is absent.
4. Refresh its existing destination, or load to the intended destination after
   checking for overlapping content. Inspect the result before saving.

See the [Power Query recovery example](powerquery.md#recovering-a-failed-create).
