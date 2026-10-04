# Excel MCP introduction video

Follow the [repository rules](../../AGENTS.md). Invoke
`/hyperframes` before changing the composition, then load the relevant skills
it selects. Framework rules, media treatments, and workflow selection belong
to those skills, not copied instructions in this project.

## Checks and preview

Run `npm run check` after every composition HTML change. Fix all errors before
handoff and review warnings before rendering.

For a Studio handoff, use HyperFrames' persistent preview, not a shell background
wrapper around the foreground server. These commands use the project's pinned
CLI through the `dev` script:

```powershell
npm run dev -- --background  # preview --background
npm run dev -- --status      # preview --status
npm run dev -- --stop        # preview --stop
```

Verify readiness with `preview --status`, keep the preview alive through review,
and stop it explicitly with `preview --stop` afterward.

## Version and render quality

`package.json` pins the HyperFrames CLI so the project remains reproducible.
Review upgrades with the unpinned latest CLI, not the project's old version:

```powershell
npx hyperframes@latest upgrade --project . --check
npx hyperframes@latest upgrade --project .
```

Validate rendered output before accepting an upgrade. Keep any required
render-profile changes with the version bump.

Production output uses `npm run render` with `--quality=high`, `--crf=18`, and
`--resolution=landscape-4k`, and writes `renders/excel-mcp-intro-4k.mp4`.
Use `npm run render:draft` only for previews. Before a final render, run the
project check, confirm the actual quality/profile flags, and record the
HyperFrames version in project metadata or commit notes.

## YouTube publication

A local render does not update YouTube. YouTube cannot replace the file behind
an existing video ID; upload the production MP4 as a new video.

1. Verify the output file's dimensions and encoding settings, not just the render
   command. Upload `renders/excel-mcp-intro-4k.mp4`, never the draft.
2. Preserve the approved title and description. Confirm upload completion,
   YouTube checks, and 4K processing. Obtain approval before publishing.
3. Verify the new public watch page and its 2160p playback option. Record the new
   video ID and YouTube publication timestamp; do not infer publication from a
   local file, an old public video, or a clean Git working tree.
4. Update `README.md`, `gh-pages/docs/index.md`, `gh-pages/overrides/main.html`,
   and `gh-pages/sitegen/sitemap.py` together, including thumbnails and publication dates.
5. Run the website checks in [gh-pages/README.md](../../gh-pages/README.md), merge
   through a PR, and verify the deployed homepage and sitemap use the new ID.
   Leave previous videos unchanged unless the user explicitly asks otherwise.
