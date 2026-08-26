# Issue tracker: GitHub

Issues and specifications for this repository live in GitHub Issues. Use the `gh` command-line tool for all operations.

## Conventions

- Infer the repository from the current Git remote.
- Pull requests are not treated as incoming work requests.
- A bare `#42` may identify an issue or pull request; check the pull request first, then the issue.
- When a skill says "publish to the issue tracker," create a GitHub issue.
- When a skill says "fetch the relevant ticket," read that GitHub issue and its comments.

## Choose the matching issue template

Read and preserve the headings from the closest template:

- General defect: `.github/ISSUE_TEMPLATE/bug_report.md`
- MCP Server defect: `.github/ISSUE_TEMPLATE/mcp_server_issue.md`
- New or changed behavior: `.github/ISSUE_TEMPLATE/feature_request.md`

Fill every section that applies. Use `N/A` when a required section does not apply, and never include private workbook data, credentials, or customer information.

`breaking-changes-issue.md` is a historical implementation plan, not a template for new issues.

## GitHub operations

- Create: `gh issue create --title "..." --body-file <path>`
- Read: `gh issue view <number> --comments --json number,title,body,labels,comments,state`
- List: `gh issue list --state open --json number,title,body,labels`
- Comment: `gh issue comment <number> --body "..."`
- Add or remove labels: `gh issue edit <number> --add-label "..."` or `--remove-label "..."`
- Close: `gh issue close <number> --comment "..."`
