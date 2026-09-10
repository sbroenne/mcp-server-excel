---
applyTo: "src/ExcelMcp.Core/Commands/Connection/**/*.cs,src/ExcelMcp.Core/Utilities/ConnectionStringSanitizer.cs,tests/**/ConnectionCommandsTests*.cs,tests/**/ConnectionTestHelper.cs,tests/**/ConnectionTestsFixture.cs"
excludeAgent: "code-review"
---

# Connection pitfalls

- `connection create` rejects `TEXT;`/`URL;`. Direct text/web imports belong to
  `querytable`; transformations and Power Query load management to `powerquery`.
  A Power Query connection's OLEDB provider does not make it an ordinary
  connection. User workflows live in `skills/shared/querytable.md` and
  `skills/shared/powerquery.md`.
- Reuse `CreateConnection`/`Connections.Add2` with command-type selection intact.
- Legacy text imports may expose TEXT or WEB access behavior; scope fallback
  to that case rather than treating the types as interchangeable.
- Cleanup must use exact WorkbookConnection/mashup identity, not a name prefix.
- `connection test` inspects configuration, not source reachability. Refresh
  status/cancellation use typed OLEDB/ODBC helpers; other types may not support it.
- Provider choice matters in tests: TEXT lifecycle tests cannot establish OLEDB
  loading behavior.
