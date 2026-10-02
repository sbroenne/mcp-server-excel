# Low-noise Awesome Copilot listing updates

This optional workflow updates **existing** `excel-cli` / `excel-mcp` listings.
It never submits a new-plugin issue, publishes a release, merges/closes a PR, asks
for upstream labels/reviews, or overwrites a human-edited branch.
Automatic submissions are disabled unless the source repository variable
`AWESOME_COPILOT_UPDATES_ENABLED` is exactly `true`.

## Repository roles and versions

| Repository | Role |
| --- | --- |
| `sbroenne/mcp-server-excel` | Canonical formatting skills, reference docs, generators, overlays, comparison, workflows and this guide |
| `sbroenne/mcp-server-excel-plugins` | Generated output only; immutable plugin tags and the commits used by listings |
| `github/awesome-copilot` | Upstream listings on `main`; receives a contribution PR |
| `sbroenne/awesome-copilot` | Writable fork; receives only the automated PR branch, based on **upstream main**, never assumed-synced fork main |

Product releases can outnumber plugin releases. A version-only product release
creates **no** plugin commit/push/tag and does not invoke this updater.
The retained plugins still launch npm packages through `npx ...@latest`.
Do not assume that a product tag exists in the plugin output repository.
See [publication rules and exact-source repair](publish-plugins-setup.md).

Source and output repositories have different commits. For example, published
tag `v2.1.0` resolves to plugin output commit
`49eac98387184790afbed75bc588282db4121b0a`; a source release SHA is not a valid
replacement for that listing pointer.

## Deterministic content rules

`PluginContent.mjs` is shared by the publisher and updater. The publisher compares
the complete prepared publication, including source-owned root overlays. The
updater compares each complete plugin directory against its **actually listed**
published commit/tag on current upstream `main`. It validates ref/SHA agreement,
manifest identities, required files, and version stamps before comparison.
Missing/invalid inputs and API data fail visibly, never as a no-op.

Only the manifest's top-level version and the exact two known version files are
ignored per plugin. Publication also ignores the corresponding generated root
marketplace entries' version fields; no other root metadata is ignored.
JSON key order/formatting is irrelevant; JSON values/array order are preserved.
Duplicate keys and numbers that would lose precision during parsing fail visibly
rather than silently comparing different values as equal.
Launch arguments, helpers, manifest metadata, documentation, skills, generated
references, assets and file additions/removals are real changes. Arbitrary
version-looking strings such as `npx package@2.3.0` or README compatibility text
are **not** stripped.

Examples:

- Product/npm `2.1.1` changes runtime internals but produces identical plugins:
  no publication and no agent or upstream PR.
- Only the publication root README changes: publish output, no listing PR.
- Only CLI skill references change: update only `excel-cli`.
- Both plugin directories change: one proposal includes both entries.
- Manifest description changes: update that listing field along with the
  published pointer. A curated upstream description is retained when the
  released manifest description did not change.
- Upstream still lists `2.0.1` after several product releases: comparison against
  that pinned commit catches all accumulated distributed changes. Comparison
  against only the immediately prior product release would miss them.

Older **existing published tags** work for catch-up, provided they do not
downgrade either a current listing or a pending proposal.

## Credentials and secure setup

Two separate credentials serve the updater:

1. `COPILOT_GITHUB_TOKEN` authenticates the Copilot engine. Its existing presence
   is not proof of usable inference or of fork/upstream write access. Validate
   it independently with the supported Copilot token requirements.
2. `AWESOME_COPILOT_PR_TOKEN` authenticates only the guarded safe-output job.
   Use an expiring **classic PAT with `public_repo`**, owned by `sbroenne`, for
   these public fork contributions. No `repo`, `workflow` or administration
   scopes are needed. The fork must remain a writable fork of upstream.

A fine-grained token restricted to the fork is **not sufficient** to create a
PR in an unrelated organization where its user is not a member.
GitHub documents this [fine-grained PAT limitation](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens#fine-grained-personal-access-tokens-limitations).
Organization PAT restrictions or SSO authorization can still block a classic
token. `public_repo` covers other accessible public repositories, not just these
two; the code's fixed repo/path allowlists are additional safeguards, not token
scope narrowing. Treat the credential accordingly.

Do not reuse `RELEASE_PAT`, `PLUGINS_REPO_TOKEN`, or the inference token based on
assumed permissions. Create/rotate credentials outside the workflow, then enter
them through the interactive secret prompt (never token values in chat, command
arguments, committed files, logs, or artifacts):

```powershell
gh secret set AWESOME_COPILOT_PR_TOKEN --repo sbroenne/mcp-server-excel
gh secret set COPILOT_GITHUB_TOKEN --repo sbroenne/mcp-server-excel
```

Set a short expiration, record ownership outside the repository, rotate before
expiry, and revoke the old token after validating the replacement. Only the
safe-output submit step sees the PR credential. Checkouts do not persist it.
Authenticated API discovery runs trusted source code only in `pre_activation`.
It uploads public snapshots bound to the published tag/commit, upstream commit,
owned PR state and a canonical digest. The separate `build` job downloads only
this public artifact, uses public Git clones to revalidate commits/listings and
the discovery digest, and runs upstream npm without any GitHub API, PR,
inference, publication or release token in the build process ancestry.
The build job requests `permissions: {}` and checkouts do not persist auth.
It also overrides gh-aw's optional global telemetry headers/endpoints to empty values.
GitHub-hosted runner services and artifact/checkout actions still use platform
transport credentials; this is not a claim that the entire runner is
credential-free. Those actions are trusted, separate completed processes,
not token-bearing parents of upstream npm. Custom application credentials and
credential artifacts never enter the build job.
The writer uses another job/runner and a fresh trusted source checkout; upstream
background processes and code changes cannot carry over. Filtering a child's
environment alone is insufficient: on Linux a same-user child can inspect its
parent's initial environment. Builds reject parents containing any known
GitHub/inference/write token, even when a sanitized child environment is supplied.
The token-bearing writer never runs upstream builds or scripts; it rechecks
public source/PR inputs and verifies the precheck's exact built file bytes.
External Git preparation (clone, fetch, checkout and content reads) and every
upstream npm install/validation/build use copied environments that exclude the
PR, GitHub, inference, publication and release credentials, including case
variants on Windows. Injected Git configuration/auth headers and credential-
helper environment variables are excluded too. Other build settings are preserved. Only trusted
GitHub API operations and the allowlisted push receive explicit authentication;
the original writer environment is not modified or treated as a safe place to
execute upstream code.
The agent runs with read-only GitHub permissions. The publisher's separate
`PLUGINS_REPO_TOKEN` continues to cover plugin output publication only.

## Preview first, then opt in

No secrets or repository settings were changed by this implementation.
Do not use a real release or an upstream test PR for verification.

Run a no-write preview of an existing output tag:

```powershell
gh workflow run update-awesome-copilot.lock.yml --repo sbroenne/mcp-server-excel -f published_tag=v2.1.0 -f preview=true
```

The deterministic precheck resolves upstream `main` and the output tag, compares
listings and any owned pending PR, and builds/validates an actionable patch in a
disposable upstream checkout. Preview does **not** start Copilot or any write
job. Its plan is retained as an Actions artifact; the summary records the
candidate output commit and comparison outcome. This validates read access and
the local patch, not the unexercised write credential or Copilot inference.
Verify those credentials separately before authorizing actual submission.

Local authenticated discovery and build must also be separate commands:

```powershell
# In a trusted shell with a read-only GH_TOKEN, discovery runs no upstream code.
$workRoot = Join-Path ([IO.Path]::GetTempPath()) "excel-listing-$([Guid]::NewGuid().ToString('N'))"
node scripts\Update-AwesomeCopilot.mjs discover v2.1.0 "$workRoot\discovery-work" "$workRoot\public-discovery.json"
# In a separately started shell whose ancestry never received application tokens:
node scripts\Update-AwesomeCopilot.mjs build "$workRoot\public-discovery.json" "$workRoot\build-work" "$workRoot\plugin-update-plan.json"
```

Do not late-unset a token or launch a filtered child from a token-bearing shell
as a substitute for the second boundary. Each work directory must be new.
Pass only the public `$workRoot` path to the separately started build shell.
Disposable workspaces must not overlap the trusted source checkout. Roots,
ancestors and leaf destinations must be ordinary directories/files: symlinks,
Windows junctions and all reparse points are rejected before reads/writes,
including dangling links or missing leaves underneath linked ancestors.
Windows local runs require `pwsh` for the complete reparse-attribute check.
Windows 8.3 and case aliases are expanded with `GetLongPathNameW` and compared
with native realpaths after every-component reparse checks. Subsequent paths and
trusted-source overlap checks use physical canonical roots, including missing
file suffixes. Legitimate short names are not treated as escapes; junctions and
other reparse points remain forbidden. POSIX retains strict lexical/realpath equality.
Git blob-mode checks also reject tracked links materialized as ordinary text
by `core.symlinks=false`. Both listing files are checked before any mutation
and again after upstream build; artifact/template/receipt paths have the same
filesystem protection. The updater does not recursively clean external trees;
remove only exact owned disposable directories without following links.
`prepare` remains an anonymous convenience command only from such a token-free
shell: public REST reads use unauthenticated `curl`, never saved `gh` login
credentials. HTTP/rate-limit/auth failures are visible; use trusted authenticated
`discover` and a separately started `build` instead of weakening isolation.

After the user validates credentials and preview:

```powershell
gh variable set AWESOME_COPILOT_UPDATES_ENABLED --repo sbroenne/mcp-server-excel --body true
```

Successful publication with real changed plugin content then calls the compiled
reusable workflow with its **actual** published tag and `preview=false`.
Root-only publications and skipped publications do not call it.
To disable writes again, set the variable to `false` (or delete it).

Independent manual catch-up:

```powershell
gh workflow run update-awesome-copilot.lock.yml --repo sbroenne/mcp-server-excel -f published_tag=v2.1.0 -f preview=false
```

This does not republish anything and can use an older retained published tag.
Actual manual writes also require the explicit opt-in. It is safe to retry a
missed/failed updater after a successful publication; no new product release or
plugin tag is required. Copilot runs only when the precheck finds an actionable
proposal. The custom job also explicitly honors `GH_AW_SAFE_OUTPUTS_STAGED=true`;
never assume gh-aw staged mode suppresses arbitrary custom scripts.

## One open PR, protected writes

Runs are serialized, with cancellation disabled. One owned open PR combines
both plugins. Ownership requires the fixed fork, `sbroenne` author, branch
`excel-plugin-updates-<12 hex characters>`, upstream `main`, and the stable
`excel-plugin-update-state` body marker; no upstream label permission is needed.

If a candidate matches the pending proposal's normalized fingerprints, no
branch, title or body is written, even if its release stamps changed. New
meaningful content refreshes the **same** PR and preserves its other proposed
plugin. The marker records the expected head, original upstream base, proposed
entries and normalized fingerprints.
Preservation applies only while the newer published tag still contains the
proposed change. If it restores a pending plugin to currently listed content,
the precheck fails visibly and leaves the PR untouched rather than reporting a
successful no-op or carrying the outdated entry into another plugin's refresh.
An authorized owner must inspect and resolve that stale proposal manually; the
workflow does not close it or remove entries automatically.
Every owned PR must also have a valid fingerprint of its visible body; removing
or corrupting that field does not bypass protection.
New state markers encode canonical UTF-8 JSON as unpadded `b64url:` base64url,
so listing text containing HTML comment delimiters cannot end the marker.
Decoding rejects malformed/noncanonical encoding, invalid UTF-8 and invalid
state or fingerprints. Original raw-JSON markers are accepted only when their
JSON is canonical, contains no double hyphen or angle brackets, and passes all
existing body/head/ownership checks. Unsafe or broken legacy markers fail
visibly and require separately authorized, verified manual migration; automation
does not reconstruct them or overwrite the PR. Writes and saved recovery bodies
use the same encoded marker.

The custom gh-aw safe output is deliberate: the built-in create handler's
non-fast-forward fallback is not this workflow's policy. The released
`v0.89.21` supports fork `head-repo` create/update and custom safe-output jobs;
the latter adds deterministic entry guards and forbids replacement PRs.
The writer requires exactly one schema-valid output, ignores agent-authored
patches, independently rechecks the trusted precheck's current public inputs and
exact built files without executing upstream code, and verifies both the
expected branch head and body immediately before writing. Only
`plugins/external.json` and generated `.github/plugin/marketplace.json` may
change, and only affected Excel entries; all other entries/order/root metadata
must remain identical. Patch size is limited to two MiB.

Branches use ordinary non-force pushes. A fresh branch starts from upstream main;
a refresh appends to the recorded PR head without rewriting history or merging
unrelated fork-main changes. A human branch/body change, conflicted upstream
listing, existing orphan branch, multiple owned open PRs, closed/merged race, or
blocked push stops visibly. No replacement PR, force push, auto-close, auto-merge,
fallback issue, failure issue, or edits to another person's PR.
Creation checks all owned-prefix fork heads, not just the current proposal's
branch. A head associated with a historical merged/closed owned PR is not an
orphan; both open and historical PR lookups are paginated.

An identical proposal closed without merge is blocked rather than resubmitted.
Upstream rejection needs a human decision. The workflow does not remove its
marker, reopen it or invent a new submission issue. If a push succeeded but
PR creation/body update failed, inspect that branch/PR manually; the next run
will protect the orphan/unrecorded head rather than overwrite it.

### Interrupted submission recovery

A refresh checks the fork's exact pushed head and unchanged upstream body while
allowing up to eight API propagation/update attempts, separated by two seconds.
It accepts only the recorded old-to-new head transition. A lost API response
can be reconciled within that same job if the next read shows the exact intended
body and head; unrelated commits, body edits or PR closure stop immediately.

The safe-output job preserves `awesome-copilot-submission-<run-id>` even on
failure. Its receipt records the old head/body fingerprint, exact pushed head,
intended body (including its state marker), allowed file contents and whether
the push completed. It contains no credentials. It is evidence, **not** authority
to accept a changed branch or to bypass the next precheck.

If all attempts fail, separately authorized human reconciliation is required:

1. Download the receipt from the authentic source workflow run, not a PR comment
   or user-provided replacement. Confirm `status=pushed`, the fixed upstream/
   fork, existing open PR number, and its owned branch/author/base.
2. Verify the fork and PR still point exactly to the receipt's new head, whose
   sole parent is its recorded old head. Compare the two allowed files byte for
   byte with the receipt and confirm no other paths changed. A different head,
   additional commit or file change is not an accepted transition.
3. Verify the complete existing PR body's SHA-256 equals `expectedBody`. If it
   changed, stop rather than overwriting human edits. If both head and body are
   unchanged as required, update **only that same PR's body** to the receipt's
   exact body through the operator's normal GitHub interface. Do not change its
   branch, title, other plugin proposal or open/closed state.
4. Run a fresh no-write catch-up preview; it must validate the new marker/head/
   body and pending published locators before any further automated submission.

For push-success/PR-create-failure, the receipt has no existing PR number. Do not
automatically create another branch: the orphan guard blocks all new proposals.
An authorized operator must inspect the recorded branch/content and decide its
disposition. Neither retries nor receipts reopen declined PRs or accept arbitrary
human commits.

## Local checks and workflow compilation

Requirements: selected .NET SDK (`global.json`), PowerShell, Node.js 22+, Git,
GitHub CLI and public read access. No desktop Excel is needed for these scripts.

```powershell
node --test --test-concurrency=3 `
    tests\ExcelMcp.SkillGeneration.Tests\PluginPublication.test.mjs `
    tests\ExcelMcp.SkillGeneration.Tests\PluginPublicationMarketplace.test.mjs `
    tests\ExcelMcp.SkillGeneration.Tests\PluginPublicationHistory.test.mjs `
    tests\ExcelMcp.SkillGeneration.Tests\PluginPublicationStaging.test.mjs
dotnet test tests\ExcelMcp.SkillGeneration.Tests\ExcelMcp.SkillGeneration.Tests.csproj -c Release --filter "Feature=PluginPublication" --blame-hang-timeout 5m
$env:PREVIEW='true'
$env:AWESOME_COPILOT_UPDATES_ENABLED='false'
$workRoot = Join-Path ([IO.Path]::GetTempPath()) "excel-listing-$([Guid]::NewGuid().ToString('N'))"
node scripts\Update-AwesomeCopilot.mjs prepare v2.1.0 "$workRoot\awesome-preview" "$workRoot\awesome-preview-plan.json"
```

The four test files share the existing checks and fixture helpers. The direct
Node command runs at most three files at once; the .NET theory runs each file
as a separate, sequential test with its own three-minute deadline. Each fixture
owns its disposable repositories and local remote; no Excel or real publication
is involved. A timeout terminates the owned process tree and includes the active
check's diagnostics. The five-minute test-host hang limit leaves time for each
test's cleanup and reporting.

Use a new disposable directory on each preview. The prepare command runs
upstream's documented `npm ci --ignore-scripts --no-audit --no-fund`,
`npm run plugin:validate`, and `npm run build`. It rejects unrelated generated
changes rather than discarding them. Existing upstream metadata warnings for
other listings may appear; failures still stop the run.

The Node regressions are also called by the repository's existing Excel-free
SkillGeneration test project, including actual publication against disposable
local Git repositories and remotes. CI exercises the suite; changed script/
workflow paths select its focused local checks without requiring COM tests.

Pin the released compiler (not an unreleased `main`) and regenerate:

```powershell
gh extension install github/gh-aw --pin v0.89.21
gh aw version
gh aw compile update-awesome-copilot --gh-aw-ref v0.89.21 --validate --no-check-update --approve
```

`--approve` records a reviewed compiler action/secret manifest; it does not grant
repository permissions, configure a credential, enable submissions or bypass Git
hooks. Review every new action/secret first. The compiled file pins gh-aw to
release commit `c35393777e5604a63721d09512263b1383301d4f`.
Commit the Markdown and generated lock workflow together; never hand-edit the
lock file. Verify the submit job depends on successful threat detection and the
agent job has no PR write credential. Detection is fail-closed; missing-tool,
missing-data, detection, incomplete-run and failure reporting cannot create issues.

## Failures and recovery

| Outcome | Action |
| --- | --- |
| No meaningful changes | Keep the existing listing/publication tag; no action needed |
| Missing plugin tag | Choose an existing **output** tag; source product tags can be intentionally absent |
| Tag/SHA/stamp/manifest mismatch | Correct source packaging or use separately authorized exact publication repair; do not substitute `main` |
| Read/API/search failure | Restore access or retry; incomplete API data is not a no-op |
| Downgrade | Choose a tag at least as new as current/pending listings |
| Inference/write authentication failure | Validate/rotate the appropriate separate credential; never expose it |
| Human/conflicting branch/body, orphan branch, declined identical proposal | Inspect manually and agree on a resolution; no automated overwrite/replacement |
| Updater failed after publication | Retry this workflow using the reported existing output tag; successful publication remains intact |
