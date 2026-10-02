---
"excelmcp": patch
---

Improve runtime package installation, workbook-session recovery, and input
validation. File preflight now reports access checks separately from workbook
validity without inspecting workbook contents. Unsupported operations fail
explicitly rather than using approximations.

Preserve underlying numeric values during formatted range copies, report
hidden worksheets correctly, and retain exact-workbook recovery ownership
after uncertain opens. Wait for desktop Excel readiness before handing off a
workbook and reject unsupported cross-workbook worksheet selection.
Mac range writes and copies now honor protected-write defaults, compatible
larger copy destinations repeat source content, and formula reads return
canonical Excel error metadata. Merged-cell writes fail before mutation
because Excel's Apple Events API cannot identify the top-left exception
reliably.

Apple Silicon macOS support is experimental beta, not Windows feature parity.
Power Query, VBA, Data Model/DAX/OLAP, Tables, PivotTables, charts, slicers,
connections, QueryTables, XML Maps, screenshots, advanced visual formatting,
and Python result reads remain unsupported. Windows retains its complete
backend. Confirmed Mac sessions now save during normal shutdown, and
concurrent close requests with conflicting save choices are rejected.
