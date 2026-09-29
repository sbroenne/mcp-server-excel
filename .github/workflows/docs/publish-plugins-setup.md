# GitHub Copilot Plugin Publishing

The source repository owns plugin templates, shared guidance, authored assets,
generation, validation, and publication. `sbroenne/mcp-server-excel-plugins` is
output-only: never fix generated files there by hand.

The two plugins contain wrappers and complete skills, not bundled runtimes.
Wrappers download the newest Windows runtime from the source repository's GitHub
Releases, checking its exact `SHA256SUMS` entry before extraction.

## Required secret

Keep the existing `PLUGINS_REPO_TOKEN` repository secret. It needs contents-write
access to `sbroenne/mcp-server-excel-plugins`; a suitably scoped PAT or GitHub App
token is sufficient. No additional cross-repository credential is required.
Create the credential outside the workflow and store it with:

```powershell
gh secret set PLUGINS_REPO_TOKEN --repo sbroenne/mcp-server-excel
```

The workflow checks that the secret exists and can reach the destination before
attempting publication. Never print the credential.

## Release handoff

`release.yml` builds and verifies every package through
`scripts\Build-ReleasePackages.ps1`. Each standalone runtime is published once;
the extension and Claude bundle reuse the MCP executable. Complete skills are
generated once for the package set and consumed by the skill ZIP, plugins, and
extension.

After GitHub Release assets exist, the release calls `publish-plugins.yml` as a
reusable workflow with the exact tag, final release commit, version, and prepared
plugin artifact name. The publisher verifies that these agree and checks out the
exact tagged source for synchronization. It never searches for a tag on the
workflow's original source commit or stamps current `main` with an older version.

Plugin publication does not depend on npm/NuGet propagation or MCP registry
registration. Marketplace publication also reports its own result. A partial
publication failure stays visible and can be repaired independently.

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
output does not create an empty commit. An absent publication tag can be repaired
without rewriting an existing tag.

## Maintenance and updates

Unsynchronized hand edits in the published repository are prohibited and may be
overwritten. For every change:

1. Edit canonical inputs under `.github/plugins/`, `skills/templates`,
   `skills/shared`, `skills/assets`, or their owning build/workflow files.
2. Build Release and generate complete skills using
   `scripts\Build-AgentSkills.ps1 -GenerateOnly`.
3. Build versioned plugins using `scripts\Build-Plugins.ps1 -Version <version>`.
   Use `-SkillsDirectory` to consume an explicit prepared skills directory.
4. Run `scripts\Sync-PublishedPluginRepo.ps1` against a disposable local output
   directory and run its generated `tests\Test-Plugins.ps1`. Inspect the complete
   publication tree, including `.github/plugin/marketplace.json`.
5. Fix failures in this source repository and regenerate, rather than patching
   output. Merge and publish only when separately authorized.

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
- **Registry failed but plugins succeeded:** repair registry registration
  separately; do not rebuild or republish already successful plugins.
