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
token is sufficient for publication. Awesome Copilot listings are edited manually;
publication does not update them.
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
direct npx configuration rather than a runtime. Complete skills are generated once for the package set and consumed by the
plugins and extension.

After GitHub Release assets and npm packages exist, the release calls
`publish-plugins.yml` as a reusable workflow with the exact tag, final release
commit, version, and prepared plugin artifact name. Waiting for npm ensures the
plugins' default npx launch path is available when the marketplace update lands.
The publisher verifies that the inputs agree and checks out the exact tagged
source for synchronization. It never searches for a tag on the workflow's
original source commit or stamps current `main` with an older version.

Marketplace publication does not depend on MCP registry registration. A partial
publication failure stays visible. Plugin publication retains its independent
repair path; registry failures are recovered through a new patch release.

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
  `marketplace.json` before migration), and in the generated Claude Code
  `.claude-plugin/marketplace.json`. Validation first requires the Claude catalog
  to list exactly those two plugins with the same `source` and `version` as the
  canonical catalog, so a stale Claude entry fails instead of being ignored.

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
but does **not** suggest an Awesome Copilot listing update.

If the normalized tree matches both baselines, publication is **skipped entirely**.
The existing plugin files, version, commit and tag remain; no new plugin tag is
created merely to match a product release. GitHub/npm product publication has
already happened and proceeds independently. For example, plugins at `v2.0.1`
can remain at that tag through product releases `v2.0.2` and `v2.1.0`. A real
plugin content change in product `v2.2.0` publishes plugins at `v2.2.0`.
Launchers still use `npx ...@latest`, so sparse plugin versions do not pin the
runtime to an older product.

The step summary and the publisher's JSON result record the decision, actual
tag/commit and changed plugins:

| Field | Meaning |
| --- | --- |
| `status` | `skipped` or `published` |
| `published_tag` | Actual retained/created published plugin tag, not necessarily the product tag |
| `published_commit` | Commit resolved from that immutable tag in the **output** repository |
| `changed_plugins` | Meaningfully changed plugin directories; empty for root-only changes |
| `handoff` | True only for a successful new publication with changed distributed plugin content |

Invalid manifests/stamps, missing tags/content, inaccessible repositories and
downgrades are failures, not a no-change outcome.

## Awesome Copilot listings are manual

Neither `release.yml` nor `publish-plugins.yml` updates Awesome Copilot
listings. A listing problem can therefore never change the release or plugin
publication result. When `handoff` is true, the step summary names the actual
`published_tag` and `published_commit` in `sbroenne/mcp-server-excel-plugins`.

1. Inspect that immutable output commit and the changed plugin directories.
2. Edit the corresponding plugin listings in `github/awesome-copilot` manually.
   Keep their source paths and pin them to the actual output commit, not a
   source-repository commit or an unpublished product tag.
3. Review the listing diff and open a normal upstream pull request. Check for an
   existing listing PR before creating another.

The automatic listing updater has been removed. Maintainers may separately
retire its unused credential and configuration; this source change does not
modify repository secrets, variables or permissions.

## Manual repair

Current skill and plugin build commands run directly in PowerShell, with no
C# packaging bootstrap. They use the source checkout containing the invoked
script. Exact-release repair still consumes the verified prepared release
payload or the explicitly requested tagged source, not current untagged content.
Preparation and publication remain separate. PR/release jobs validate generated
skills early; publication also validates downloaded or recovered skills.

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
manual listing updates must use an **existing published tag**, not every product tag.
The next automatic release also compares its candidate with that retained tag,
not only repaired `main`. If a content-changing repair is already present on
`main`, a later otherwise-version-only release creates a new immutable plugin
tag that makes the repair reachable by listings. Changed-plugin handoff is
computed against the old immutable tag; destination staging still uses `HEAD`.
Restoring `main` to content already distributed by the retained tag does not
trigger a plugin listing update. Invalid immutable baseline content fails
visibly rather than allowing a false no-change result.

### v2.5.0 checkout failure

Release run `39304106786`, plugin job `118243153092`, checked out source commit
`655d82dc2f43b492de9e32bb6ffca3af28a97860` for `v2.5.0`. The logged
`d4fa6e9fce2acbb66c67b411ce21370f73ccb8b5` was the output repository's existing
commit, not the source release commit. Publication stopped at the source
checkout guard, before changing the output repository.

The old error combined an unexpected source commit and uncommitted source
changes without reporting either value. The job log therefore does not establish
what caused the rejection. Fresh local checkouts of this exact release passed
the guard, including with the job's Git `2.56.0.windows.2`. Do not assume a
credentials or line-ending problem, reset the source, or disable the guard.
The guard now reports an unexpected checkout commit separately from Git's
changed-path list, without printing file contents or credentials.

An authorized manual repair on unchanged `main`, run `39310269423`, repeated
the original rejection. A subsequent authorized exact-release repair using the
diagnostic branch, run `39310607520`, succeeded with the split checks unchanged
in strictness. It published output commit
`974bd11c9c3d87263cc5887078bbcbf94ab05f09` and immutable tag `v2.5.0`, with
both marketplace plugin versions set to `2.5.0`. The reason the original
combined expression rejected the hosted checkout remains unproven; no
source-changing cause was established.

The v2.5.0 plugin repair is complete. Do not repeat it or the other release
jobs just to test these diagnostics. If a future repair is necessary, obtain
separate publication authorization and use the plugin-only workflow with an
existing release tag. For example, the exact-v2.5.0 command is:

```powershell
gh workflow run publish-plugins.yml --ref main -f release_tag=v2.5.0
```

This starts a new manual run, not a rerun of the original release jobs. After
the diagnostic change is merged, it uses the updated publication guard from
`main`, resolves the existing `v2.5.0` source tag, and
downloads only that release's `excel-plugins-v2.5.0.zip` and `SHA256SUMS`.
The recorded payload checksum is
`7efd2a67892e5467adfa20e2f33f6a06674be3136e183eaf557d7c32b7a09bb6`;
the workflow must still verify the downloaded checksum entry itself.
It does not republish GitHub release assets, npm, NuGet, VS Code or the MCP
registry, and does not rewrite existing tags. No new product release is needed
for this plugin repair.

Inspect the new job's source commit and result. If it rejects the checkout,
retain the reported commit/changed paths and stop: fix the identified cause in
the source repository before another authorized attempt. If it succeeds, verify
the reported output tag/commit and plugin contents in the output repository.
An Awesome Copilot listing update remains a separate, explicitly authorized
operation.

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
   publication tree, including `.github/plugin/marketplace.json` and
   `.claude-plugin/marketplace.json`.
5. Fix failures in this source repository and regenerate, rather than patching
   output. Merge and publish only when separately authorized.

`Publish-PreparedPlugins.ps1` is the current source-owned guard. Preparation
continues to use the exact requested source's synchronization script/overlay,
including on repair of an older release. Both old and new source release tags
must be available to identify previously owned overlay paths.

Merging a source PR does not publish plugins. A release or authorized manual
repair must run the publication path.

Skill validation shares the deterministic layout rules used by the publication
guard. It accepts the CLI discovery-plus-formatting pair, formatting-only
packages, and immutable legacy packages with their general skill and references.
Both CLI skills must carry the exact release version. Mixed old reference
collections with the new discovery skill, unknown skill directories, and mixed
MCP layouts are rejected.

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
- **Registry failed but plugins succeeded:** fix the registry failure and create
  a new patch release. Do not manually replay the already successful plugin
  publication; normal release publication skips unchanged plugin content.
