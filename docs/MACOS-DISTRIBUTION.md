# macOS distribution readiness

ExcelMcp publishes self-contained CLI and MCP Server artifacts for `osx-arm64`.
The same target is used by GitHub Release ZIPs,
platform-targeted VSIX packages, Copilot plugin bootstrap downloads, and Claude
Desktop MCPB bundles.
The npm launchers select the `darwin-arm64` runtime package; canonical Copilot
plugins invoke those launchers through
`npx`. Agent Skills contain guidance only and remain architecture-neutral.
No launcher may silently fall back to a runtime for the wrong architecture.

## Verification status

| Target | Build | Package inspection | Clean-install launch | Excel E2E release gate |
| --- | --- | --- | --- | --- |
| macOS arm64 | Native | Mach-O architecture, executable mode, code signature, archive contents | Verified on Apple Silicon | Prompt-free CLI and MCP workflows run on the serialized Mac/Excel runner |
| Windows x64 | Cross-published or native | PE header and archive contents | Not executed on macOS | Covered separately on Windows with Excel |

Intel Macs are unsupported. Launchers, plugins, and extension runtime selection
must fail closed instead of selecting the Apple Silicon runtime on x64 macOS.

## Signing and notarization

`scripts/Sign-MacBinary.ps1` always gives packaged Mach-O executables a
verifiable signature. Without release secrets it uses an ad-hoc signature and
prints an explicit warning. With a configured Developer ID Application
certificate, release jobs use hardened runtime signing and timestamping.

`scripts/Submit-MacNotarization.ps1` submits release ZIP, MCPB, and VSIX
payloads only when all App Store Connect API credentials and a Developer ID
identity are configured. Nonstandard archive extensions are submitted through
a temporary ZIP copy accepted by `notarytool`. Missing credentials produce an
explicit unnotarized result; partial credentials or an unsigned submission fail
the build. These archives cannot be stapled, so acceptance is established from
the successful `notarytool` submission result.

Required release secrets:

- `MACOS_CERTIFICATE_P12`
- `MACOS_CERTIFICATE_PASSWORD`
- `MACOS_SIGNING_IDENTITY`
- `MACOS_NOTARY_API_KEY_P8`
- `MACOS_NOTARY_KEY_ID`
- `MACOS_NOTARY_ISSUER_ID`

Checksums protect transport integrity independently of Apple signing status.
They must not be described as proof of Developer ID signing or notarization.

## Local package checks

On Apple Silicon macOS, build packages and inspect them with:

```powershell
./scripts/Test-DistributionPackages.ps1 `
  -ArchivePath ./path/to/macos-arm64.zip `
  -ExecutableRelativePath mcp-excel `
  -ExpectedArchitecture macos-arm64 `
  -Launch

./scripts/Test-DistributionPackages.ps1 `
  -ArchivePath ./path/to/windows.zip `
  -ExecutableRelativePath mcp-excel.exe `
  -ExpectedArchitecture windows-x64
```

The launch switch is intentionally restricted to the host's native Apple
Silicon architecture. Windows artifacts receive structure-only inspection.
