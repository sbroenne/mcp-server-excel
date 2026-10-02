# Issue tracker: GitHub

Work requests live in this repository's GitHub Issues; PRs are not incoming
work requests.

## Conventions

- A bare `#42` may identify an issue or pull request; check the pull request first, then the issue.
- When a skill says "publish to the issue tracker," create a GitHub issue.
- When a skill says "fetch the relevant ticket," read that GitHub issue and its comments.

## Choose the matching issue template

Choose the closest issue form. When creating an issue through a tool rather
than the website, render each form field as a heading and preserve its order:

- General defect: `.github/ISSUE_TEMPLATE/bug_report.yml`
- MCP Server defect: `.github/ISSUE_TEMPLATE/mcp_server_issue.yml`
- New or changed behavior: `.github/ISSUE_TEMPLATE/feature_request.yml`

Use `N/A` when a required section does not apply.

The [pre-1.0 breaking-changes plan](history/breaking-changes-pre-1.0.md) is
historical, not a template or current API guidance. Report vulnerabilities
privately through GitHub Security Advisories, never as public work requests.
