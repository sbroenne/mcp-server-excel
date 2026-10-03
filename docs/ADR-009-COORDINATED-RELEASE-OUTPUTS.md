# ADR-009: Coordinate releases and publish prepared outputs

**Status:** Current

## Context and decision

The release workflow owns component version metadata and changelog generation,
then prepares the distributions from the release inputs. Plugin publication
consumes prepared output and compares content before writing. Unchanged plugin
content can retain its existing tag even when the runtime release advances.

## Reasons and tradeoffs

Independent component versioning and hand-edited package copies would let
documentation, launchers, skills, and binaries drift. Coordinated preparation
keeps release metadata and output ownership explicit.

Publishing every time regardless of content would create meaningless plugin
commits and tags. Comparing prepared content avoids that churn, but means a
plugin tag is not necessarily the latest runtime version.

Coordinated releases couple packaging validation across components. A source
edit or local build is not publication; distribution updates still depend on
their release and publishing workflows.

## Implementation and guidance

- [Release workflow](../.github/workflows/release.yml)
- [Prepared plugin publication](../scripts/Publish-PreparedPlugins.ps1)
- [Release procedures](RELEASE-STRATEGY.md)
- [Plugin maintenance](../.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates)
