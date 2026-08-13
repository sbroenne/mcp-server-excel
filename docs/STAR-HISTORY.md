# Star History Data

The Pages workflow renders the repository's star-history SVG from aggregate
`date,count` rows. It never stores GitHub usernames, user IDs, or stargazer node
IDs.

## Data sources

The committed `.github/star-history.csv` bootstrap was generated with a
maintainer-authenticated GraphQL query that selected only
`repository.stargazers.edges.starredAt`. The timestamps were immediately grouped
by UTC date into cumulative counts, and the final aggregate count was checked
against the repository's public `stargazers_count`.

Scheduled builds do not query the restricted timestamped stargazers REST
endpoint. They read the exact current `stargazers_count` from public repository
metadata using the workflow's `GITHUB_TOKEN`. Before recording the count, the
workflow restores the CSV from `star-history-data`; an expected missing branch
or file uses the committed aggregate bootstrap, while other API failures stop
the deployment. It then appends or replaces that UTC day's aggregate snapshot
and persists the CSV on the data branch. Count decreases are retained so
unstars are represented rather than hidden.

The dedicated data branch keeps daily state durable without granting a personal
access token or write access to `main`. The Pages site continues to deploy from
the standard GitHub Pages artifact.

## Validation

Run the aggregate parser and SVG regression checks from the repository root:

```powershell
./scripts/Test-StarHistory.ps1
```
