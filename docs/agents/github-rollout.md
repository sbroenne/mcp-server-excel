# GitHub configuration rollout

Repository-file changes do not authorize live settings changes, publication,
merging, or credential migration. Re-read settings immediately before any
separately approved rollout. The read-only baseline below was inspected on
2026-10-02; it is not a settings write payload.

## Merge protection

| Baseline | Proposed change after successful PR runs and approval |
| --- | --- |
| Required `CI Gate`, `Docs Site`, `dependency-review` | Retain all three; add `CodeQL Completion` and `Verify changeset present` only after GitHub exposes their exact check names |
| Zero approving reviews, squash only | Preserve the single-maintainer policy |
| Copilot re-review on pushes, no draft reviews | Preserve existing behavior |
| Legacy protection requires conversation resolution and prevents force pushes/deletion | Keep it until the replacement ruleset explicitly preserves those protections |
| Administrator bypass permits release metadata writes | Do not tighten before replacing and validating the release writer |
| Up-to-date checks disabled; update-branch button disabled | Consider both together after verifying release automation compatibility |

Bind required checks to GitHub Actions where supported. Test code changes,
documentation-only changes, forks, and `skip-changelog` before enforcement.
Require the stable CodeQL completion result, not all four language jobs:
selective PR scanning intentionally omits unchanged languages. Code scanning
alert thresholds are distinct from successful execution/upload checks.
Keep merge queues out of this rollout; all required workflows would need
`merge_group` support first.

## Workflow permissions and credentials

Authored action references are commit-pinned; generated agentic-workflow locks
remain compiler-owned. Pinning does not itself restrict allowed actions.
Keep read-only default tokens and job-scoped write grants.

| Privileged workflow | Existing use to preserve when designing policies |
| --- | --- |
| `release.yml` | Authorized manual release; `RELEASE_PAT` writes metadata to main; tag/release writes and package publishing have separate jobs |
| `publish-plugins.yml` | Exact-release publication to the output-only plugin repository; `PLUGINS_REPO_TOKEN`; never runs listing updates |
| `update-awesome-copilot.lock.yml` | Manual-only (`workflow_dispatch`), compiled, guarded discovery/proposal/write stages with separate inference and PR credentials; explicit opt-in |
| `publish-mcp-registry.yml` | Exact-release registration and `mcp-registry` environment with OIDC |
| `deploy-gh-pages.yml` | Read-only dependency/build job; separate Pages/OIDC deployment job |
| `usage-analytics.yml` | Azure OIDC collection, sanitized agent input, separate branch writer; no writes on fork PRs |
| `star-history.yml` | Public aggregate collection and separate branch writer; no writes on PRs |
| `link-check.yml` | Scheduled/manual external link check with issue creation; not a required PR gate |
| `dependency-review.yml` | PR dependency/license enforcement and summary comments |

Before tightening execution policies, inventory exact actors/events, nested
reusable-workflow permissions, and permitted writers. GitHub's policy
**Evaluate** mode requires Enterprise Cloud and is not promised for this
personal repository. Preserve existing legitimate PR writers before changing
Actions' create/approve permission. No credential replacement, rotation, or
secret-value inspection is part of this change.

OIDC still uses the default name-based subject. Migrate external Azure/registry
trust first; only then separately approve immutable-subject adoption.
The old Actions `copilot` environment is not the current agent secret store;
use dedicated **Agents** secrets/variables.

## Releases

Release immutability was disabled in the inspected baseline. Before enabling
it for future releases, merge and verify draft preparation, complete asset
digest checks, non-destructive replay, and explicit old-mutable repair behavior
from [release strategy](../RELEASE-STRATEGY.md#github-asset-integrity-and-replay).
Do not dispatch release/publishing workflows as tests.

`RELEASE-INPUTS.json` records the workflow path/original checkout, exact metadata
patch, final release commit/tag, and artifact digests. It is not an SBOM or signed
attestation. Provenance attestation remains deferred until the patched build
can be represented and verified honestly. Do not rewrite old tags or claim
the final release commit was the untouched original build checkout.

## Availability and unresolved verification

- GitHub Code Quality explicitly reported unavailable for this repository.
  Keep the working advanced CodeQL setup; do not enable competing default setup.
- Hosted cloud automations require private/internal repositories and are
  excluded for this public repository.
- Review effort and Copilot approval switches remain unverified. Leave them
  unchanged. Current documented choices are Lite and Balanced; Max is not yet
  available. Balanced and agent evaluation can incur additional costs.
- Secret-results merge protection eligibility is not established. Free public
  secret scanning does not establish entitlement to every preview rule.
- Windows cloud-agent firewall/internet and repository MCP settings remain
  unverified and unchanged. Hosted runners do not contain Excel.
- App UI-only project instructions and trust state could not be read through
  the available session API. Verify the native review-and-accept flow locally;
  do not assume YAML parsing proves that it was exercised.
- The inspected baseline used a standalone VS Code review checklist. Current
  shared instruction discovery and review guidance are documented in
  [agent development](development.md). Verify loaded files in the actual
  clients; their presence alone does not establish adherence.

The inspected main branch had 16 dependency alerts in `llm-tests/uv.lock`.
This change selects PyJWT 2.15.1 and urllib3 2.8.0, outside the reported
vulnerable ranges. One PyJWT advisory did not specify a first-patched version;
scanner re-evaluation after delivery is still required. Local lockfile changes
do not close alerts on main.
The pre-existing test-fixture HTML-filter finding is not addressed by workflow
hardening. Related unpinned-action findings are addressed in authored workflow
files, but their live scanner state must be rechecked after delivery.

## Delivery boundary

Record exact commands and results in the PR template; internal/configuration
changes use `skip-changelog`. Ask before committing, pushing, or creating a PR
for this work. Merging, publishing, and live settings changes require separate
authorization even after file-change approval.
