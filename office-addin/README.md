# ExcelMcp Office.js bridge

This package is an optional, versioned foundation for future macOS Excel
capability tiers. It does not enable any workbook feature action. The base
Apple Events backend remains usable when the add-in is absent.

The bridge binds only to `127.0.0.1`, requires HTTPS, authenticates every API
request, validates task-pane origins, binds sessions to exact workbook URLs,
serializes requests per workbook, and expires requests at their deadline.

See [macOS Office.js setup](../docs/MACOS-OFFICEJS.md) for installation,
activation, health, upgrade, and removal commands.
