# Cheritto Decision Inbox — Workflow Warnings Cleanup

## 2026-04-25 — Publish Plugins tag resolution must use git checkout state

- **Decision:** `publish-plugins.yml` now checks out the source repo with tags and resolves the release tag with `git tag --points-at ${{ github.event.workflow_run.head_sha }}` (with a short retry), instead of querying `git/matching-refs/tags/v` and comparing `.object.sha`.
- **Why:** This repo creates annotated release tags. The REST matching-refs payload returns the tag object's SHA, not the target commit SHA, so the old comparison could never find the release tag for a successful `workflow_run`.
- **Impact:** The follow-on plugin publish workflow can now resolve the released version reliably after `release.yml` completes, and the maintainer docs were aligned to describe the cross-repo token requirements accurately (PAT or app token).

## 2026-04-25 — Remove retired Azure workflow files, keep only the reference docs

- **Decision:** Delete `.github/workflows/deploy-azure-runner.yml.disabled` and `.github/workflows/integration-tests.yml.disabled`.
- **Why:** They were no longer executable workflows, and the only remaining value was as a historical re-enable hint. Keeping the stale files in-tree made the repo look like those flows were merely paused, while the real state is that Azure self-hosted Excel CI has been retired.
- **Impact:** Existing docs were minimally updated to stop referencing `.disabled` filenames and instead describe the Azure runner guide as historical infrastructure reference. `publish-plugins` did **not** fall into the same bucket: it required an actual workflow fix now, not just monitoring.
