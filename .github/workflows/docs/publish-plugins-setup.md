# GitHub Copilot Plugin Publishing

The source repository owns plugin templates, shared guidance, authored assets,
generation, validation, and publication. `sbroenne/mcp-server-excel-plugins` is
output-only: never fix generated files there by hand.

The two plugins contain launch configuration, an argument-safe CLI wrapper, and
complete skills, not bundled runtimes. They use the public npm packages through
`npx` and require Node.js 18 or later.

## Required secret

Keep the existing `PLUGINS_REPO_TOKEN` repository secret. It needs contents-write
access to `sbroenne/mcp-server-excel-plugins`; a suitably scoped PAT or GitHub App
token is sufficient for publication. The optional Awesome Copilot updater has
separate credentials; the publishing token is never assumed to authorize it.
Create the credential outside the workflow and store it with:

```powershell
gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel
```

The workflow checks that the secret exists and can reach the destination before
attempting publication. Never print the credential.

## Release handoff

`release.yml` builds and verifies every package through
`scripts\Build-ReleasePackages.ps1`. Each standalone runtime is published once;
the extension reuses the MCP executable, while the Claude bundle contains a
direct npx configuration rather than a runtime. Complete skills are
generated once for the package set and consumed by the skill ZIP, plugins, and
extension.

After GitHub Release assets and npm packages exist, the release calls
`publish-plugins.yml` as a reusable workflow with the exact tag, final release
commit, version, and prepared plugin artifact name. Waiting for npm ensures the
plugins' default npx launch path is available when the marketplace update lands.
The publisher verifies that the inputs agree and checks out the exact tagged
source for synchronization. It never searches for a tag on the workflow's
original source commit or stamps current `main` with an older version.

Marketplace publication does not depend on MCP registry registration. A partial
publication failure stays visible and can be repaired independently.

## Publish only changed output

After validating the prepared plugins and skills, the publisher prepares a
disposable complete publication tree from the exact source release. It compares
the files Git will distribute against both publication `HEAD` and its current
immutable published tag **before** staging the destination, committing, pushing,
or creating a tag. Changed source
paths are not sufficient: generated references can change with service contracts.

The comparison ignores only:

- Each `plugins/<name>/plugin.json` **top-level** `version`, after validating the
  manifest identity, schema and version.
- Exactly `plugins/<name>/version.txt` and
  `plugins/<name>/skills/<name>-report-formatting/VERSION`, after verifying their
  release stamps. Legacy immutable baselines retain their old `<name>/VERSION`
  path for comparison; mixed skill layouts are rejected.
- The `version` of the corresponding `excel-cli` and `excel-mcp` entries in the
  generated root `.github/plugin/marketplace.json` (or the validated legacy
  `marketplace.json` before migration).

JSON objects are compared canonically, including launch JSON; object formatting
and key order do not matter, but array order and meaningful values do. Other root
metadata, nested versions, launch arguments (including pinned npm versions),
helpers, README files, skills, references, assets, modes and added/removed files
all matter. No general version-number stripping or documentation exclusions.
Ordinary Git text clean conversion determines distributed bytes, not a blanket
documentation/newline exclusion.
Destination staging uses the same conversion rules, forces only the exact
prepared/removed paths, and verifies its staged Git tree equals the validated
candidate before any commit or tag. Legitimate ignored-name files are included;
unrelated ignored local files are not staged.
Known candidate files are explicitly re-read during staging, even when copying
preserves a cached timestamp and a changed version/content has the same size.
The output checkout retains the matching text-conversion setting so a later
Git file-stat refresh does not report generated CRLF files as uncommitted changes.
File/directory replacements remove only empty directories left by removed
tracked files. A remaining nonempty destination directory blocks publication
rather than deleting ignored or otherwise unrelated local files.

Root overlay files under `.github/plugins/marketplace-repo` are included.
Previously source-owned overlay files are removed from the candidate when
removed in the exact released source. Unowned root files such as the existing
license are retained and compared. A root-overlay-only change publishes output
but does **not** request an Awesome Copilot listing update.

If the normalized tree matches both baselines, publication is **skipped entirely**.
The existing plugin files, version, commit and tag remain; no new plugin tag is
created merely to match a product release. GitHub/npm product publication has
already happened and proceeds independently. For example, plugins at `v2.0.1`
can remain at that tag through product releases `v2.0.2` and `v2.1.0`. A real
plugin content change in product `v2.2.0` publishes plugins at `v2.2.0`.
Launchers still use `npx ...@latest`, so sparse plugin versions do not pin the
runtime to an older product.

Reusable workflow outputs:

| Output | Meaning |
| --- | --- |
| `status` | `skipped` or `published` |
| `published_tag` | Actual retained/created published plugin tag, not necessarily the product tag |
| `published_commit` | Commit resolved from that immutable tag in the **output** repository |
| `changed_plugins` | JSON array of meaningfully changed plugin directories; empty for root-only changes |
| `handoff` | True only for a successful new publication with changed distributed plugin content |

The step summary records the decision, actual tag/commit and changed plugins.
Invalid manifests/stamps, missing tags/content, inaccessible repositories and
downgrades are failures, not a no-change outcome.

Only a true handoff and `AWESOME_COPILOT_UPDATES_ENABLED=true` call the optional
updater. Its failure is visible in its own job and can be retried independently;
it does not undo plugin or product publication. See
[Awesome Copilot setup](awesome-copilot-update-setup.md).

Both reusable-workflow caller jobs (`release.yml`'s `publish-plugins` and
`publish-plugins.yml`'s `update-awesome-copilot`) must grant `actions: read`,
`contents: read`, and `pull-requests: read` for the compiled updater's jobs.
GitHub validates nested permissions before evaluating the optional job's `if`,
even with the opt-in disabled. These job-scoped read grants do not enable updates
or change the publisher's default `contents: read` permissions.

## Manual repair

With publication authorization, rerun an existing release:

```powershell
gh workflow run publish-plugins.yml -f release_tag=v1.2.3
```

The workflow downloads that release's `excel-plugins-v1.2.3.zip` and verifies its
exact `SHA256SUMS` entry before extracting it. For older
releases without the prepared payload, it builds the exact tagged source using
that source's generator. Missing or unsupported legacy inputs fail; there is no
fallback to current source.

The workflow serializes publication, blocks downgrades and mismatched tags,
validates plugin identities and wrapper-only contents, and validates both skills
with the pinned official Agent Skills validator. Automatic duplicates are
skipped. Manual repair compares generated output before committing; identical
output does not create an empty commit. The explicit manual path uses raw output
differences too, so it can restore missing/mismatched known version stamps and
create an absent exact publication tag even when normalized content is unchanged.
Malformed manifests, identities and other missing required content still fail
visibly. It does not substitute current source or bypass downgrade validation.

An existing immutable tag is never rewritten. A repair can update publication
`main`, but `published_commit` still describes the old immutable tag. Repairs of
existing tags do not automatically hand off a different `main` commit as that
tag. Repairing publication and catching up a listing are distinct tasks:
the updater accepts an **existing published tag**, not every product tag.
The next automatic release also compares its candidate with that retained tag,
not only repaired `main`. If a content-changing repair is already present on
`main`, a later otherwise-version-only release creates a new immutable plugin
tag that makes the repair reachable by listings. Changed-plugin handoff is
computed against the old immutable tag; destination staging still uses `HEAD`.
Restoring `main` to content already distributed by the retained tag does not
trigger a plugin listing update. Invalid immutable baseline content fails
visibly rather than allowing a false no-change result.

## Maintenance and updates

Unsynchronized hand edits in the published repository are prohibited and may be
overwritten. For every change:

1. Edit canonical inputs under `.github/plugins/`, the actual `skills/<name>`
   directories, `docs/reference/report-formatting.md`, or their owning
   build/workflow files. General reference docs are not packaged into plugins.
2. Build Release and generate complete skills using
   `scripts\Build-AgentSkills.ps1 -GenerateOnly`.
3. Build versioned plugins using `scripts\Build-Plugins.ps1 -Version <version>`.
   Use `-SkillsDirectory` to consume an explicit prepared skills directory.
4. Run `scripts\Sync-PublishedPluginRepo.ps1` against a disposable local output
   directory and run its generated `tests\Test-Plugins.ps1`. Inspect the complete
   publication tree, including `.github/plugin/marketplace.json`.
5. Fix failures in this source repository and regenerate, rather than patching
   output. Merge and publish only when separately authorized.

`Publish-PreparedPlugins.ps1` is the current source-owned guard. Preparation
continues to use the exact requested source's synchronization script/overlay,
including on repair of an older release. Both old and new source release tags
must be available to identify previously owned overlay paths.

Merging a source PR does not publish plugins. A release or authorized manual
repair must run the publication path.

## Local validation without publishing

From the source repository root:

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release
.\scripts\Build-AgentSkills.ps1 -GenerateOnly
.\scripts\Build-Plugins.ps1 -Version 1.2.3 -OutputDir artifacts\plugin-check
New-Item -ItemType Directory artifacts\publication-check
.\scripts\Sync-PublishedPluginRepo.ps1 -PublishedRepoDir artifacts\publication-check -BuiltPluginsDir artifacts\plugin-check -Version 1.2.3
& .\artifacts\publication-check\tests\Test-Plugins.ps1
```

These commands create local output only. Do not dispatch a real release as a test.

## Troubleshooting

- **Missing token or inaccessible target:** configure or rotate
  `PLUGINS_REPO_TOKEN`; keep permissions limited to the target repository.
- **Tag/commit/version mismatch:** use the intended existing release, not its
  original pre-metadata workflow commit.
- **Downgrade blocked:** an older release cannot overwrite newer publication.
- **Missing skill input:** build Release, then explicitly generate skills.
- **No commit:** the destination already matches the prepared output.
- **Publication skipped:** only validated release bookkeeping changed; use the
  reported retained tag for an independent marketplace catch-up, not the newer
  product tag.
- **Missing current publication tag:** use authorized exact-release manual repair;
  normal publication will not silently invent a baseline.
- **Registry failed but plugins succeeded:** repair registry registration
  separately; do not rebuild or republish already successful plugins.
