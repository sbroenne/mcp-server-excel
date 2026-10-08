# Build and release constraints

- Keep workflow SDK setup compatible with `global.json`. Preserve analyzer and
  warning-as-error settings in `Directory.Build.props` and `.editorconfig`.
- CLI pre-build service stopping remains enabled on self-hosted desktops even
  when `CI=true`; live local daemons can otherwise lock rebuild output. Hosted CI
  skips it. Do not add temporary cleanup-client builds or graceful shutdown waits.
- `ci.yml` has Excel-free runtime and documentation gates. Local pre-commit
  selects checks by changed paths; preserve runtime/non-runtime and merge-parent
  handling so imported changes do not trigger unrelated Excel E2E.
- `Get-ValidationPlan.ps1` is the shared CI, CodeQL, and local selection source.
  Distinguish shipped documentation from developer instructions, and keep each
  tooling project's filter separate. Preserve binary-package dependencies;
  runtime edits must not automatically select unrelated publication checks.
  Keep required completion checks always reporting; detection failure,
  cancellation, and unexpected skips must fail them. Main/manual CI and
  main/merge-group/scheduled/manual CodeQL retain complete coverage.
  C# CodeQL keeps traced compilation for generated code.
  `Build-CiInputs.ps1` builds selected test projects, but source/count checks
  still require a complete Release solution build. Cache dependency downloads,
  not writable compiled outputs. Packages build their own required binaries.
  Hosted test partitions use separate checkouts so rebuild tests cannot race
  packaging. Selection and local Excel group commands: `tests/README.md`.
- Local CLI builds invoke `scripts\Stop-ExcelCliService.ps1` once. Forcibly stop
  only background CLI services from this worktree's output, validating PID/start
  time. Builds do not save workbooks or terminate Excel, MCP, foreground CLI, or
  other worktrees' services. Test-run cleanup passes an explicit private pipe.
  Preserve ordinary product save/close and explicit service-stop safeguards.
- `release.yml` owns versions and changelog generation. Do not dispatch it as a
  test. Authorized merges use squash.

Procedures: `docs/RELEASE-STRATEGY.md`, `vscode-extension/DEVELOPMENT.md`, and
`.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates`.
