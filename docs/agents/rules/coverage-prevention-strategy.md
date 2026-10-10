# Generated contract checks

After changing a Core contract or generator, build Release and inspect generated
CLI options, batch JSON dispatch, Service routing, and MCP schemas. Names,
aliases, defaults, validation, results, and timeouts must agree. Hand-written MCP
tools may still own atomic no-session behavior, cancellation, or extra metadata.

Run `Invoke-ExcelFreeTests.ps1 -Local -Contracts` and
`check-doc-counts.ps1 -SkipBuild` under
`scripts` after that build. Compilation alone does not detect missing routes.

Adding, renaming, or removing an MCP tool action, CLI command, or CLI category
also requires updating `.github/usage-analytics-weights.json`: give every action
an effort level, homepage feature, and (for CLI categories) the matching MCP
tools. `UsageAnalyticsWeightsTests` in the McpServer and CLI test projects fail
when the file and code disagree; update the file rather than the tests. Choose
levels with the rules in
[Public Usage Analytics Automation](../../DEVELOPMENT.md#public-usage-analytics-automation).
